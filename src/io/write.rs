use std::collections::{HashMap, HashSet};
use std::fs::{File, rename};
use std::io;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use chrono::NaiveDateTime;
use zip::write::FileOptions;

use crate::{SCell, Sheet, StyleFor, StyleOrigin, StyleUse, ucell, Value, ValueFormat, ValueType, WorkBook};
use crate::attrmap::AttrMap;
use crate::error::OdsError;
use crate::format::FormatPartType;
use crate::io::tmp2zip::{TempWrite, TempZip};
use crate::io::xmlwriter::XmlWriter;
use crate::refs::{CellRange, cellranges_string};
use crate::style::{FontFaceDecl, HeaderFooter, PageLayout, Style};
use crate::text::{TextElem, TextVec};

// this did not work out as expected ...
// TODO: find out why this breaks content.xml
// type OdsWriter = zip::ZipWriter<BufWriter<File>>;
// type XmlOdsWriter<'a> = quick_xml::Writer<&'a mut zip::ZipWriter<BufWriter<File>>>;

type OdsWriter = TempZip;
type XmlOdsWriter<'a> = XmlWriter<TempWrite<'a>>;

/// Writes the ODS file.
pub fn write_ods<P: AsRef<Path>>(book: &WorkBook, ods_path: P) -> Result<(), OdsError> {
    write_ods_flags(book, ods_path, false, true, true)?;
    Ok(())
}

/// Writes the ODS file. The parameter clean indicates the cleanup of the
/// temp files at the end.
pub fn write_ods_flags<P: AsRef<Path>>(book: &WorkBook,
                                       ods_path: P,
                                       bak: bool,
                                       zip: bool,
                                       clean: bool) -> Result<(), OdsError> {
    if bak && ods_path.as_ref().exists() {
        let mut ods_bak = ods_path.as_ref().to_path_buf();
        ods_bak.set_extension("bak");
        rename(&ods_path, &ods_bak)?;
    }

    // Origin File if any
    let orig = if let Some(file) = &book.file {
        if !file.exists() {
            let mut ods_bak = ods_path.as_ref().to_path_buf();
            ods_bak.set_extension("bak");
            Some(ods_bak)
        } else {
            Some(file.clone())
        }
    } else {
        None
    };

    // let zip_file = File::create(ods_path)?;
    // let mut zip_writer = zip::ZipWriter::new(io::BufWriter::new(zip_file));
    let mut zip_writer = TempZip::new(ods_path.as_ref());

    let mut file_set = HashSet::<String>::new();
    //
    if let Some(orig) = orig {
        copy_workbook(&orig, &mut file_set, &mut zip_writer)?;
    }

    write_mimetype(&mut zip_writer, &mut file_set)?;
    write_manifest(&mut zip_writer, &mut file_set)?;
    write_manifest_rdf(&mut zip_writer, &mut file_set)?;
    write_meta(&mut zip_writer, &mut file_set)?;
    //write_settings(&mut zip_writer, &mut file_set)?;
    //write_configurations(&mut zip_writer, &mut file_set)?;
    write_ods_styles(&book, &mut zip_writer, &mut file_set)?;
    write_ods_content(&book, &mut zip_writer, &mut file_set)?;

    if zip {
        zip_writer.zip()?;
    }
    if clean {
        zip_writer.clean()?;
    }

    Ok(())
}

fn copy_workbook(ods_orig_name: &PathBuf, file_set: &mut HashSet<String>, zip_writer: &mut OdsWriter) -> Result<(), OdsError> {
    let ods_orig = File::open(ods_orig_name)?;
    let mut zip_orig = zip::ZipArchive::new(ods_orig)?;

    for i in 0..zip_orig.len() {
        let mut zip_entry = zip_orig.by_index(i)?;

        if zip_entry.is_dir() {
            if !file_set.contains(zip_entry.name()) {
                file_set.insert(zip_entry.name().to_string());
                zip_writer.add_directory(zip_entry.name(), FileOptions::default())?;
            }
        } else if !file_set.contains(zip_entry.name()) {
            file_set.insert(zip_entry.name().to_string());
            let mut wr = zip_writer.start_file(zip_entry.name(), FileOptions::default())?;
            let mut buf = [0u8; 1024];
            loop {
                let n = zip_entry.read(&mut buf)?;
                if n == 0 {
                    break;
                } else {
                    wr.write_all(&buf[0..n])?;
                }
            }
        }
    }

    Ok(())
}

fn write_mimetype(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), io::Error> {
    if !file_set.contains("mimetype") {
        file_set.insert(String::from("mimetype"));

        let mut w = zip_out.start_file("mimetype", FileOptions::default().compression_method(zip::CompressionMethod::Stored))?;

        let mime = "application/vnd.oasis.opendocument.spreadsheet";
        w.write_all(mime.as_bytes())?;
    }

    Ok(())
}

fn write_manifest(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("META-INF/manifest.xml") {
        file_set.insert(String::from("META-INF/manifest.xml"));

        zip_out.add_directory("META-INF", FileOptions::default())?;
        let w = zip_out.start_file("META-INF/manifest.xml", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("manifest:manifest")?;
        xml_out.attr("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")?;
        xml_out.attr("manifest:version", "1.2")?;

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "/")?;
        xml_out.attr("manifest:version", "1.2")?;
        xml_out.attr("manifest:media-type", "application/vnd.oasis.opendocument.spreadsheet")?;

//        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
//            ("manifest:full-path", String::from("Configurations2/")),
//            ("manifest:media-type", String::from("application/vnd.sun.xml.ui.configuration")),
//        ]))?;

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "manifest.rdf")?;
        xml_out.attr("manifest:media-type", "application/rdf+xml")?;

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "styles.xml")?;
        xml_out.attr("manifest:media-type", "text/xml")?;

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "meta.xml")?;
        xml_out.attr("manifest:media-type", "text/xml")?;

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "content.xml")?;
        xml_out.attr("manifest:media-type", "text/xml")?;

//        xml_out.write_event(xml::xml_empty_a("manifest:file-entry", vec![
//            ("manifest:full-path", String::from("settings.xml")),
//            ("manifest:media-type", String::from("text/xml")),
//        ]))?;

        xml_out.end_elem("manifest:manifest")?;

        xml_out.close()?;
    }

    Ok(())
}

fn write_manifest_rdf(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("manifest.rdf") {
        file_set.insert(String::from("manifest.rdf"));

        let w = zip_out.start_file("manifest.rdf", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("rdf:RDF")?;
        xml_out.attr("xmlns:rdf", "http://www.w3.org/1999/02/22-rdf-syntax-ns#")?;

        xml_out.elem("rdf:Description")?;
        xml_out.attr("rdf:about", "content.xml")?;

        xml_out.empty("rdf:type")?;
        xml_out.attr("rdf:resource", "http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile")?;

        xml_out.end_elem("rdf:Description")?;

        xml_out.elem("rdf:Description")?;
        xml_out.attr("rdf:about", "")?;

        xml_out.empty("ns0:hasPart")?;
        xml_out.attr("xmlns:ns0", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#")?;
        xml_out.attr("rdf:resource", "content.xml")?;

        xml_out.end_elem("rdf:Description")?;

        xml_out.elem("rdf:Description")?;
        xml_out.attr("rdf:about", "")?;

        xml_out.empty("rdf:type")?;
        xml_out.attr("rdf:resource", "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document")?;

        xml_out.end_elem("rdf:Description")?;

        xml_out.end_elem("rdf:RDF")?;

        xml_out.close()?;
    }

    Ok(())
}

fn write_meta(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("meta.xml") {
        file_set.insert(String::from("meta.xml"));

        let w = zip_out.start_file("meta.xml", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("office:document-meta")?;
        xml_out.attr("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")?;
        xml_out.attr("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")?;
        xml_out.attr("office:version", "1.2")?;

        xml_out.elem("office:meta")?;

        xml_out.elem_text("meta:generator", "spreadsheet-ods 0.1.0")?;
        let s = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?;
        let d = NaiveDateTime::from_timestamp(s.as_secs() as i64, 0);
        xml_out.elem_text("meta:creation-date", &d.format("%Y-%m-%dT%H:%M:%S%.f").to_string())?;
        xml_out.elem_text("meta:editing-duration", "P0D")?;
        xml_out.elem_text("meta:editing-cycles", "1")?;
        // xml_out.elem_text_esc("meta:initial-creator", &username::get_user_name().unwrap())?;

        // TODO: allow to set this data.

        xml_out.end_elem("office:meta")?;

        xml_out.end_elem("office:document-meta")?;

        xml_out.close()?;
    }

    Ok(())
}

//fn write_settings(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
//    if !file_set.contains("settings.xml") {
//        file_set.insert(String::from("settings.xml"));
//
//        zip_out.start_file("settings.xml", FileOptions::default())?;
//
//        let mut xml_out = XmlWriter::new(zip_out);
//
//        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
//        xml_out.write(b"\n")?;
//
//        xml_out.write_event(xml_start_a("office:document-settings", vec![
//            ("xmlns:office", String::from("urn:oasis:names:tc:opendocument:xmlns:office:1.0")),
//            ("xmlns:ooo", String::from("http://openoffice.org/2004/office")),
//            ("xmlns:config", String::from("urn:oasis:names:tc:opendocument:xmlns:config:1.0")),
//            ("office:version", String::from("1.2")),
//        ]))?;
//
//        xml_out.write_event(xml_start("office:settings"))?;
//        xml_out.write_event(xml::xml_end("office:settings"))?;
//
//        xml_out.write_event(xml::xml_end("office:document-settings"))?;
//    }
//
//    Ok(())
//}

//fn write_configurations(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
//    if !file_set.contains("Configurations2") {
//        file_set.insert(String::from("Configurations2"));
//        file_set.insert(String::from("Configurations2/accelerator"));
//        file_set.insert(String::from("Configurations2/floater"));
//        file_set.insert(String::from("Configurations2/images"));
//        file_set.insert(String::from("Configurations2/images/Bitmaps"));
//        file_set.insert(String::from("Configurations2/menubar"));
//        file_set.insert(String::from("Configurations2/popupmenu"));
//        file_set.insert(String::from("Configurations2/progressbar"));
//        file_set.insert(String::from("Configurations2/statusbar"));
//        file_set.insert(String::from("Configurations2/toolbar"));
//        file_set.insert(String::from("Configurations2/toolpanel"));
//
//        zip_out.add_directory("Configurations2", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/accelerator", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/floater", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/images", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/images/Bitmaps", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/menubar", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/popupmenu", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/progressbar", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/statusbar", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/toolbar", FileOptions::default())?;
//        zip_out.add_directory("Configurations2/toolpanel", FileOptions::default())?;
//    }
//
//    Ok(())
//}

fn write_ods_styles(book: &WorkBook,
                    zip_out: &mut OdsWriter,
                    file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("styles.xml") {
        file_set.insert(String::from("styles.xml"));

        let w = zip_out.start_file("styles.xml", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("office:document-styles")?;
        xml_out.attr("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")?;
        xml_out.attr("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")?;
        xml_out.attr("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")?;
        xml_out.attr("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")?;
        xml_out.attr("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")?;
        xml_out.attr("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")?;
        xml_out.attr("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")?;
        xml_out.attr("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")?;
        xml_out.attr("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")?;
        xml_out.attr("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")?;
        xml_out.attr("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")?;
        xml_out.attr("xmlns:calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0")?;
        xml_out.attr("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")?;
        xml_out.attr("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")?;
        xml_out.attr("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0")?;
        xml_out.attr("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0")?;
        xml_out.attr("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")?;
        xml_out.attr("office:version", "1.2")?;

        xml_out.elem("office:font-face-decls")?;
        write_font_decl(&book.fonts, StyleOrigin::Styles, &mut xml_out)?;
        xml_out.end_elem("office:font-face-decls")?;

        xml_out.elem("office:styles")?;
        write_styles(&book.styles, StyleOrigin::Styles, StyleUse::Default, &mut xml_out)?;
        write_styles(&book.styles, StyleOrigin::Styles, StyleUse::Named, &mut xml_out)?;
        write_value_styles(&book.formats, StyleOrigin::Styles, StyleUse::Named, &mut xml_out)?;
        xml_out.end_elem("office:styles")?;

        xml_out.elem("office:automatic-styles")?;
        write_pagelayout(&book.page_layouts, &mut xml_out)?;
        write_styles(&book.styles, StyleOrigin::Styles, StyleUse::Automatic, &mut xml_out)?;
        write_value_styles(&book.formats, StyleOrigin::Styles, StyleUse::Automatic, &mut xml_out)?;
        xml_out.end_elem("office:automatic-styles")?;

        xml_out.elem("office:master-styles")?;
        write_masterpage(&book.page_layouts, &mut xml_out)?;
        xml_out.end_elem("office:master-styles")?;

        xml_out.end_elem("office:document-styles")?;

        xml_out.close()?;
    }

    Ok(())
}

fn write_ods_content(book: &WorkBook, zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    file_set.insert(String::from("content.xml"));

    let w = zip_out.start_file("content.xml", FileOptions::default())?;
    let mut xml_out = XmlWriter::new(w);

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-content")?;
    xml_out.attr("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")?;
    xml_out.attr("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")?;
    xml_out.attr("xmlns:xhtml", "http://www.w3.org/1999/xhtml")?;
    xml_out.attr("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")?;
    xml_out.attr("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")?;
    xml_out.attr("xmlns:xforms", "http://www.w3.org/2002/xforms")?;
    xml_out.attr("xmlns:dom", "http://www.w3.org/2001/xml-events")?;
    xml_out.attr("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0")?;
    xml_out.attr("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0")?;
    xml_out.attr("xmlns:math", "http://www.w3.org/1998/Math/MathML")?;
    xml_out.attr("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")?;
    xml_out.attr("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")?;
    xml_out.attr("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")?;
    xml_out.attr("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")?;
    xml_out.attr("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")?;
    xml_out.attr("xmlns:ooo", "http://openoffice.org/2004/office")?;
    xml_out.attr("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")?;
    xml_out.attr("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")?;
    xml_out.attr("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")?;
    xml_out.attr("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")?;
    xml_out.attr("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")?;
    xml_out.attr("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")?;
    xml_out.attr("xmlns:xlink", "http://www.w3.org/1999/xlink")?;
    xml_out.attr("xmlns:formx", "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0")?;
    xml_out.attr("xmlns:dc", "http://purl.org/dc/elements/1.1/")?;
    xml_out.attr("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")?;
    xml_out.attr("xmlns:rpt", "http://openoffice.org/2005/report")?;
    xml_out.attr("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")?;
    xml_out.attr("xmlns:css3t", "http://www.w3.org/TR/css3-text/")?;
    xml_out.attr("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")?;
    xml_out.attr("xmlns:ooow", "http://openoffice.org/2004/writer")?;
    xml_out.attr("xmlns:oooc", "http://openoffice.org/2004/calc")?;
    xml_out.attr("xmlns:tableooo", "http://openoffice.org/2009/table")?;
    xml_out.attr("xmlns:calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0")?;
    xml_out.attr("xmlns:drawooo", "http://openoffice.org/2010/draw")?;
    xml_out.attr("office:version", "1.2")?;

    xml_out.empty("office:scripts")?;

    xml_out.elem("office:font-face-decls")?;
    write_font_decl(&book.fonts, StyleOrigin::Content, &mut xml_out)?;
    xml_out.end_elem("office:font-face-decls")?;

    xml_out.elem("office:automatic-styles")?;
    write_styles(&book.styles, StyleOrigin::Content, StyleUse::Automatic, &mut xml_out)?;
    write_value_styles(&book.formats, StyleOrigin::Content, StyleUse::Automatic, &mut xml_out)?;
    xml_out.end_elem("office:automatic-styles")?;

    xml_out.elem("office:body")?;
    xml_out.elem("office:spreadsheet")?;

    for sheet in &book.sheets {
        write_sheet(&book, &sheet, &mut xml_out)?;
    }

    xml_out.end_elem("office:spreadsheet")?;
    xml_out.end_elem("office:body")?;
    xml_out.end_elem("office:document-content")?;

    xml_out.close()?;

    Ok(())
}

/// Is the cell hidden, and if yes how many more columns are hit.
fn check_hidden(ranges: &[CellRange], row: ucell, col: ucell) -> (bool, ucell) {
    if let Some(found) = ranges.iter().find(|s| s.contains(row, col)) {
        (true, found.to.col - col)
    } else {
        (false, 0)
    }
}

/// Removes any outlived Ranges from the vector.
pub(crate) fn remove_outlooped(ranges: &mut Vec<CellRange>, row: ucell, col: ucell) {
    *ranges = ranges.drain(..).filter(|s| !s.out_looped(row, col)).collect();
}

fn write_sheet(book: &WorkBook, sheet: &Sheet, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    xml_out.elem("table:table")?;
    xml_out.attr_esc("table:name", &*sheet.name)?;
    if let Some(style) = &sheet.style {
        xml_out.attr_esc("table:style-name", style.as_str())?;
    }
    if let Some(print_ranges) = &sheet.print_ranges {
        xml_out.attr_esc("table:print-ranges", &cellranges_string(print_ranges))?;
    }

    let max_cell = sheet.used_grid_size();

    write_table_columns(&sheet, max_cell, xml_out)?;

    // list of current spans
    let mut spans = Vec::<CellRange>::new();

    // table-row + table-cell
    let mut first_cell = true;
    let mut last_r: ucell = 0;
    let mut last_c: ucell = 0;

    for ((cur_row, cur_col), cell) in sheet.data.iter() {

        // There may be a lot of gaps of any kind in our data.
        // In the XML format there is no cell identification, every gap
        // must be filled with empty rows/columns. For this we need some
        // calculations.

        // For the repeat-counter we need to look forward.
        // Works nicely with the range operator :-)
        let (next_r, next_c, is_last_cell) =
            if let Some(((next_r, next_c), _)) = sheet.data.range((*cur_row, cur_col + 1)..).next() {
                (*next_r, *next_c, false)
            } else {
                (max_cell.0, max_cell.1, true)
            };

        // Looking forward row-wise.
        let forward_dr = next_r - *cur_row;

        // Column deltas are only relevant in the same row, but we need to
        // fill up to max used columns.
        let forward_dc = if forward_dr >= 1 {
            max_cell.1 - *cur_col
        } else {
            next_c - *cur_col
        };

        // Looking backward row-wise.
        let backward_dr = *cur_row - last_r;
        // When a row changes our delta is from zero to cur_col.
        let backward_dc = if backward_dr >= 1 {
            *cur_col
        } else {
            *cur_col - last_c
        };

        // println!("cell first={} {},{} < *{},{}* < {},{} ", first_cell, last_r, last_c, cur_row, cur_col, next_r, next_c);
        // println!("     backward {},{} forward {},{}", backward_dr, backward_dc, forward_dr, forward_dc);

        // After the first cell there is always an open row tag that
        // needs to be closed.
        if backward_dr > 0 && !first_cell {
            write_end_last_row(sheet, *cur_row, backward_dr, xml_out)?;
        }

        // Any empty rows before this one?
        if backward_dr > 0 {
            write_empty_rows_before(sheet, *cur_row, first_cell, backward_dr, max_cell, xml_out)?;
        }

        // Start a new row if there is a delta or we are at the start.
        // Fills in any blank cells before the current cell.
        if backward_dr > 0 || first_cell {
            write_start_current_row(sheet, *cur_row, backward_dc, xml_out)?;
        }

        // Remove no longer usefull cell-spans.
        remove_outlooped(&mut spans, *cur_row, *cur_col);

        // Current cell is hidden?
        let (is_hidden, hidden_cols) = check_hidden(&spans, *cur_row, *cur_col);

        // And now to something completely different ...
        write_cell(book, cell, is_hidden, xml_out)?;

        // There may be some blank cells until the next one, but only one less the forward.
        if forward_dc > 1 {
            write_empty_cells(forward_dc, hidden_cols, xml_out)?;
        }

        // The last cell we will write? We can close the last row here,
        // where we have all the data.
        if is_last_cell {
            write_end_current_row(sheet, *cur_row, xml_out)?;
        }

        // maybe span. only if visible, that nicely eliminates all
        // double hides.
        if !is_hidden && (cell.span.0 > 1 || cell.span.1 > 1) {
            spans.push(CellRange::origin_span(*cur_row, *cur_col, cell.span));
        }

        first_cell = false;
        last_r = *cur_row;
        last_c = *cur_col;
    }

    xml_out.end_elem("table:table")?;

    Ok(())
}

fn write_empty_cells(mut forward_dc: u32,
                     hidden_cols: u32,
                     xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {

    // split between hidden and regular cells.
    if hidden_cols >= forward_dc {
        xml_out.empty("covered-table-cell")?;
        let repeat = (forward_dc - 1).to_string();
        xml_out.attr("table:number-columns-repeated", repeat.as_str())?;

        forward_dc = 0;
    } else if hidden_cols > 0 {
        xml_out.empty("covered-table-cell")?;
        let repeat = hidden_cols.to_string();
        xml_out.attr("table:number-columns-repeated", repeat.as_str())?;

        forward_dc -= hidden_cols;
    }

    if forward_dc > 0 {
        xml_out.empty("table:table-cell")?;
        let repeat = (forward_dc - 1).to_string();
        xml_out.attr("table:number-columns-repeated", repeat.as_str())?;
    }

    Ok(())
}

fn write_start_current_row(sheet: &Sheet,
                           cur_row: ucell,
                           backward_dc: u32,
                           xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {

    // Start of headers
    if let Some(header_rows) = &sheet.header_rows {
        if header_rows.from == cur_row {
            xml_out.elem("table:table-header-rows")?;
        }
    }

    xml_out.elem("table:table-row")?;
    if let Some(row_style) = sheet.row_style(cur_row) {
        xml_out.attr_esc("table:style-name", row_style.as_str())?;
    }

    // Might not be the first column in this row.
    if backward_dc > 0 {
        let backward_dc = backward_dc.to_string();
        xml_out.empty("table:table-cell")?;
        xml_out.attr("table:number-columns-repeated", backward_dc.as_str())?;
    }

    Ok(())
}

fn write_end_last_row(sheet: &Sheet,
                      cur_row: u32,
                      backward_dr: u32,
                      xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    xml_out.end_elem("table:table-row")?;

    // This row was the end of the header.
    if let Some(header_rows) = &sheet.header_rows {
        let last_row = cur_row - backward_dr;
        if header_rows.to == last_row {
            xml_out.end_elem("table:table-header-rows")?;
        }
    }

    Ok(())
}

fn write_end_current_row(sheet: &Sheet,
                         cur_row: u32,
                         xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    xml_out.end_elem("table:table-row")?;


    // This row was the end of the header.
    if let Some(header_rows) = &sheet.header_rows {
        if header_rows.to == cur_row {
            xml_out.end_elem("table:table-header-rows")?;
        }
    }

    Ok(())
}

fn write_empty_rows_before(sheet: &Sheet,
                           cur_row: ucell,
                           first_cell: bool,
                           mut backward_dr: u32,
                           max_cell: (ucell, ucell),
                           xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    // Empty rows in between are 1 less than the delta, except at the very start.
    let mut corr = if first_cell {
        0u32
    } else {
        1u32
    };

    // Only deltas greater 1 are relevant.
    // Or if this is the very start.
    if backward_dr > 1 || first_cell && backward_dr > 0 {

        // split up the empty rows, if there is some header stuff.
        if let Some(header_rows) = &sheet.header_rows {

            // What was the last_row? Was there a header start since?
            let last_row = cur_row - backward_dr;
            if header_rows.from < cur_row && header_rows.from > last_row {
                write_empty_row(header_rows.from - last_row - corr, max_cell, xml_out)?;
                xml_out.elem("table:table-header-rows")?;
                // Don't write the empty line for the first header-row, we can
                // collapse it with the rest. corr suits fine for this.
                corr = 0;
                // We correct the empty line count.
                backward_dr = cur_row - header_rows.from;
            }

            // What was the last row here? Was there a header end since?
            let last_row = cur_row - backward_dr;
            if header_rows.to < cur_row && header_rows.to > cur_row - backward_dr {
                // Empty lines, including the current line that marks
                // the end of the header.
                write_empty_row(header_rows.to - last_row - corr + 1, max_cell, xml_out)?;
                xml_out.end_elem("table:table-header-rows")?;
                // Correction for table start is no longer needed.
                corr = 1;
                // We correct the empty line count.
                backward_dr = cur_row - header_rows.to;
            }
        }

        // Write out the empty lines.
        write_empty_row(backward_dr - corr, max_cell, xml_out)?;
    }

    Ok(())
}

fn write_empty_row(empty_count: u32,
                   max_cell: (u32, u32),
                   xml_out: &mut XmlWriter<TempWrite>) -> Result<(), OdsError> {
    xml_out.elem("table:table-row")?;
    xml_out.attr("table:number-rows-repeated", &empty_count.to_string())?;

    // We fill the empty spaces completely up to max columns.
    let max_cell_col = max_cell.1.to_string();
    xml_out.empty("table:table-cell")?;
    xml_out.attr("table:number-columns-repeated", max_cell_col.as_str())?;

    xml_out.end_elem("table:table-row")?;

    Ok(())
}

fn write_table_columns(sheet: &Sheet,
                       max_cell: (ucell, ucell),
                       xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {

    // table:table-column
    for c in 0..max_cell.1 {
        let style = sheet.column_style(c);
        let cell_style = sheet.column_cell_style(c);

        // markup header columns
        if let Some(header_cols) = &sheet.header_cols {
            if header_cols.from == c {
                xml_out.elem("table:table-header-columns")?;
            }
        }

        if style.is_some() || cell_style.is_some() {
            xml_out.empty("table:table-column")?;
            if let Some(style) = style {
                xml_out.attr_esc("table:style-name", style.as_str())?;
            }
            if let Some(cell_style) = cell_style {
                xml_out.attr_esc("table:default-cell-style-name", cell_style.as_str())?;
            }
        } else {
            xml_out.empty("table:table-column")?;
        }

        // markup header columns
        if let Some(header_cols) = &sheet.header_cols {
            if header_cols.to == c {
                xml_out.end_elem("table:table-header-columns")?;
            }
        }
    }

    Ok(())
}

fn write_cell(book: &WorkBook,
              cell: &SCell,
              is_hidden: bool,
              xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    let tag = if is_hidden { "table:covered-table-cell" } else { "table:table-cell" };

    match cell.value {
        Value::Empty => xml_out.empty(tag)?,
        _ => xml_out.elem(tag)?,
    }

    if let Some(formula) = &cell.formula {
        xml_out.attr_esc("table:formula", formula.as_str())?;
    }

    // Direct style oder value based default style.
    if let Some(style) = &cell.style {
        xml_out.attr_esc("table:style-name", style.as_str())?;
    } else if let Some(style) = book.def_style(cell.value.value_type()) {
        xml_out.attr_esc("table:style-name", style.as_str())?;
    }

    // Spans
    if cell.span.0 > 1 {
        xml_out.attr_esc("table:number-rows-spanned", cell.span.0.to_string().as_str())?;
    }
    if cell.span.1 > 1 {
        xml_out.attr_esc("table:number-columns-spanned", cell.span.1.to_string().as_str())?;
    }

    // Might not yield a useful result. Could not exist, or be in styles.xml
    // which I don't read. Defaulting to to_string() seems reasonable.
    let value_style = if let Some(style_name) = &cell.style {
        book.find_value_format(style_name)
    } else {
        None
    };

    match &cell.value {
        Value::Empty => {}
        Value::Text(s) => {
            xml_out.attr("office:value-type", "string")?;
            for l in s.split('\n') {
                xml_out.elem("text:p")?;
                xml_out.text_esc(l)?;
                xml_out.end_elem("text:p")?;
            }
        }
        Value::TextM(t) => {
            xml_out.elem("text:p")?;
            write_textvec(t, xml_out)?;
            xml_out.end_elem("text:p")?;
        }
        Value::DateTime(d) => {
            xml_out.attr("office:value-type", "date")?;
            let value = d.format("%Y-%m-%dT%H:%M:%S%.f").to_string();
            xml_out.attr("office:date-value", value.as_str())?;
            xml_out.elem("text:p")?;
            if let Some(value_style) = value_style {
                xml_out.text_esc(value_style.format_datetime(d).as_str())?;
            } else {
                xml_out.text(d.format("%d.%m.%Y").to_string().as_str())?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::TimeDuration(d) => {
            xml_out.attr("office:value-type", "time")?;

            let mut value = String::from("PT");
            value.push_str(&d.num_hours().to_string());
            value.push_str("H");
            value.push_str(&(d.num_minutes() % 60).to_string());
            value.push_str("M");
            value.push_str(&(d.num_seconds() % 60).to_string());
            value.push_str(".");
            value.push_str(&(d.num_milliseconds() % 1000).to_string());
            value.push_str("S");

            xml_out.attr("office:time-value", value.as_str())?;

            xml_out.elem("text:p")?;
            if let Some(value_style) = value_style {
                xml_out.text_esc(value_style.format_time_duration(d).as_str())?;
            } else {
                xml_out.text(&d.num_hours().to_string())?;
                xml_out.text(":")?;
                xml_out.text(&(d.num_minutes() % 60).to_string())?;
                xml_out.text(":")?;
                xml_out.text(&(d.num_seconds() % 60).to_string())?;
                xml_out.text(".")?;
                xml_out.text(&(d.num_milliseconds() % 1000).to_string())?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::Boolean(b) => {
            xml_out.attr("office:value-type", "boolean")?;
            xml_out.attr("office:boolean-value", if *b { "true" } else { "false" })?;
            xml_out.elem("text:p")?;
            if let Some(value_style) = value_style {
                xml_out.text_esc(value_style.format_boolean(*b).as_str())?;
            } else {
                xml_out.text(if *b { "true" } else { "false" })?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::Currency(c, v) => {
            xml_out.attr("office:value-type", "currency")?;
            xml_out.attr_esc("office:currency", c.as_str())?;
            let value = v.to_string();
            xml_out.attr("office:value", value.as_str())?;
            xml_out.elem("text:p")?;
            if let Some(value_style) = value_style {
                xml_out.text_esc(value_style.format_float(*v).as_str())?;
            } else {
                xml_out.text(c)?;
                xml_out.text(" ")?;
                xml_out.text(&value)?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::Number(v) => {
            xml_out.attr("office:value-type", "float")?;
            let value = v.to_string();
            xml_out.attr("office:value", value.as_str())?;
            xml_out.elem("text:p")?;
            if let Some(value_style) = value_style {
                xml_out.text_esc(value_style.format_float(*v).as_str())?;
            } else {
                xml_out.text(value.as_str())?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::Percentage(v) => {
            xml_out.attr("office:value-type", "percentage")?;
            xml_out.attr("office:value", format!("{}%", v).as_str())?;
            xml_out.elem("text:p")?;
            if let Some(value_style) = value_style {
                xml_out.text_esc(value_style.format_float(*v * 100.0).as_str())?;
            } else {
                xml_out.text(&(v * 100.0).to_string())?;
            }
            xml_out.end_elem("text:p")?;
        }
    }

    match cell.value {
        Value::Empty => {}
        _ => xml_out.end_elem(tag)?
    }

    Ok(())
}

fn write_font_decl(fonts: &HashMap<String, FontFaceDecl>, origin: StyleOrigin, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for font in fonts.values().filter(|s| s.origin() == origin) {
        xml_out.empty("style:font-face")?;
        xml_out.attr_esc("style:name", font.name().as_str())?;
        for (a, v) in font.attr_iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    Ok(())
}

fn write_styles(styles: &HashMap<String, Style>,
                origin: StyleOrigin,
                styleuse: StyleUse,
                xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for style in styles.values().filter(|s| s.origin() == origin && s.styleuse() == styleuse) {
        if styleuse == StyleUse::Default {
            xml_out.elem("style:default-style")?;
        } else {
            xml_out.elem("style:style")?;
            xml_out.attr_esc("style:name", style.name().as_str())?;
        }
        let family = match style.family() {
            StyleFor::Table => "table",
            StyleFor::TableColumn => "table-column",
            StyleFor::TableRow => "table-row",
            StyleFor::TableCell => "table-cell",
            StyleFor::None => "",
        };
        xml_out.attr("style:family", family)?;
        if let Some(display_name) = &style.display_name() {
            xml_out.attr_esc("style:display-name", display_name.as_str())?;
        }
        if let Some(parent) = &style.parent() {
            xml_out.attr_esc("style:parent-style-name", parent.as_str())?;
        }
        if let Some(value_format) = &style.value_format() {
            xml_out.attr_esc("style:data-style-name", value_format.as_str())?;
        }

        if !style.cell().is_empty() {
            xml_out.empty("style:table-cell-properties")?;
            for (a, v) in style.cell() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
        if !style.col().is_empty() {
            xml_out.empty("style:table-column-properties")?;
            for (a, v) in style.col() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
        if !style.row().is_empty() {
            xml_out.empty("style:table-row-properties")?;
            for (a, v) in style.row() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
        if !style.table().is_empty() {
            xml_out.empty("style:table-properties")?;
            for (a, v) in style.table() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
        if !&style.paragraph().is_empty() {
            xml_out.empty("style:paragraph-properties")?;
            for (a, v) in style.paragraph() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
        if !style.text().is_empty() {
            xml_out.empty("style:text-properties")?;
            for (a, v) in style.text() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }

        if let Some(stylemaps) = style.stylemaps() {
            for sm in stylemaps {
                xml_out.empty("style:map")?;
                xml_out.attr_esc("style:condition", sm.condition())?;
                xml_out.attr_esc("style:apply-style-name", sm.applied_style())?;
                xml_out.attr_esc("style:base-cell-address", &sm.base_cell().to_string())?;
            }
        }

        if styleuse == StyleUse::Default {
            xml_out.end_elem("style:default-style")?;
        } else {
            xml_out.end_elem("style:style")?;
        }
    }

    Ok(())
}

fn write_value_styles(styles: &HashMap<String, ValueFormat>,
                      origin: StyleOrigin,
                      styleuse: StyleUse,
                      xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for style in styles.values().filter(|s| s.origin() == origin && s.styleuse() == styleuse) {
        let tag = match style.value_type() {
            ValueType::Empty => "number:empty_style", // ???
            ValueType::Boolean => "number:boolean-style",
            ValueType::Number => "number:number-style",
            ValueType::Text => "number:text-style",
            ValueType::TextM => "number:text-style",
            ValueType::TimeDuration => "number:time-style",
            ValueType::Percentage => "number:percentage-style",
            ValueType::Currency => "number:currency-style",
            ValueType::DateTime => "number:date-style",
        };

        xml_out.elem(tag)?;
        xml_out.attr_esc("style:name", style.name().as_str())?;
        if let Some(prp) = style.attr_map() {
            for (a, v) in prp.iter() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }

        if let Some(parts) = style.parts() {
            for part in parts {
                let part_tag = match part.part_type() {
                    FormatPartType::Boolean => "number:boolean",
                    FormatPartType::Number => "number:number",
                    FormatPartType::Scientific => "number:scientific-number",
                    FormatPartType::CurrencySymbol => "number:currency-symbol",
                    FormatPartType::Day => "number:day",
                    FormatPartType::Month => "number:month",
                    FormatPartType::Year => "number:year",
                    FormatPartType::Era => "number:era",
                    FormatPartType::DayOfWeek => "number:day-of-week",
                    FormatPartType::WeekOfYear => "number:week-of-year",
                    FormatPartType::Quarter => "number:quarter",
                    FormatPartType::Hours => "number:hours",
                    FormatPartType::Minutes => "number:minutes",
                    FormatPartType::Seconds => "number:seconds",
                    FormatPartType::Fraction => "number:fraction",
                    FormatPartType::AmPm => "number:am-pm",
                    FormatPartType::EmbeddedText => "number:embedded-text",
                    FormatPartType::Text => "number:text",
                    FormatPartType::TextContent => "number:text-content",
                    FormatPartType::StyleText => "style:text",
                    FormatPartType::StyleMap => "style:map",
                };

                if part.part_type() == FormatPartType::Text || part.part_type() == FormatPartType::CurrencySymbol {
                    xml_out.elem(part_tag)?;
                    if let Some(prp) = part.attr_map() {
                        for (a, v) in prp.iter() {
                            xml_out.attr_esc(a.as_ref(), v.as_str())?;
                        }
                    }
                    if let Some(content) = part.content() {
                        xml_out.text_esc(content)?;
                    }
                    xml_out.end_elem(part_tag)?;
                } else {
                    xml_out.empty(part_tag)?;
                    if let Some(prp) = part.attr_map() {
                        for (a, v) in prp.iter() {
                            xml_out.attr_esc(a.as_ref(), v.as_str())?;
                        }
                    }
                }
            }
        }

        xml_out.end_elem(tag)?;
    }

    Ok(())
}

fn write_pagelayout(styles: &HashMap<String, PageLayout>,
                    xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for style in styles.values() {
        xml_out.elem("style:page-layout")?;
        xml_out.attr_esc("style:name", &style.name())?;

        if let Some(attr) = style.attr_map() {
            xml_out.empty("style:page-layout-properties")?;
            for (k, v) in attr.iter() {
                xml_out.attr(k.as_ref(), v.as_str())?;
            }
        }

        xml_out.elem("style:header-style")?;
        xml_out.empty("style:header-footer-properties")?;
        if !style.header_attr().is_empty() {
            for (k, v) in style.header_attr() {
                xml_out.attr(k.as_ref(), v.as_str())?;
            }
        }
        xml_out.end_elem("style:header-style")?;

        xml_out.elem("style:footer-style")?;
        xml_out.empty("style:header-footer-properties")?;
        if !style.header_attr().is_empty() {
            for (k, v) in style.footer_attr() {
                xml_out.attr(k.as_ref(), v.as_str())?;
            }
        }
        xml_out.end_elem("style:footer-style")?;

        xml_out.end_elem("style:page-layout")?;
    }

    Ok(())
}

fn write_masterpage<'a>(styles: &'a HashMap<String, PageLayout>,
                        xml_out: &mut XmlOdsWriter<'a>) -> Result<(), OdsError> {
    for style in styles.values() {
        xml_out.elem("style:master-page")?;
        xml_out.attr("style:name", &style.masterpage_name())?;
        xml_out.attr("style:page-layout-name", &style.name())?;

        xml_out.elem("style:header")?;
        if !style.header().display() {
            xml_out.attr("style:display", "false")?;
        }
        write_regions(&style.header(), xml_out)?;
        xml_out.end_elem("style:header")?;

        xml_out.elem("style:header_left")?;
        if !style.header_left().display() {
            xml_out.attr("style:display", "false")?;
        }
        write_regions(&style.header_left(), xml_out)?;
        xml_out.end_elem("style:header_left")?;

        xml_out.elem("style:footer")?;
        if !style.footer().display() {
            xml_out.attr("style:display", "false")?;
        }
        write_regions(&style.footer(), xml_out)?;
        xml_out.end_elem("style:footer")?;

        xml_out.elem("style:footer_left")?;
        if !style.footer_left().display() {
            xml_out.attr("style:display", "false")?;
        }
        write_regions(&style.footer_left(), xml_out)?;
        xml_out.end_elem("style:footer_left")?;

        xml_out.end_elem("style:master-page")?;
    }

    Ok(())
}

fn write_regions<'a>(hf: &'a HeaderFooter,
                     xml_out: &mut XmlOdsWriter<'a>) -> Result<(), OdsError> {
    if !hf.left().is_empty() {
        xml_out.elem("style:region-left")?;
        xml_out.elem("text:p")?;
        write_textvec(&hf.left(), xml_out)?;
        xml_out.end_elem("text:p")?;
        xml_out.end_elem("style:region-left")?;
    }
    if !hf.center().is_empty() {
        xml_out.elem("style:region-center")?;
        xml_out.elem("text:p")?;
        write_textvec(&hf.center(), xml_out)?;
        xml_out.end_elem("text:p")?;
        xml_out.end_elem("style:region-center")?;
    }
    if !hf.right().is_empty() {
        xml_out.elem("style:region-right")?;
        xml_out.elem("text:p")?;
        write_textvec(&hf.right(), xml_out)?;
        xml_out.end_elem("text:p")?;
        xml_out.end_elem("style:region-right")?;
    }
    if !hf.content().is_empty() {
        xml_out.elem("text:p")?;
        write_textvec(&hf.content(), xml_out)?;
        xml_out.end_elem("text:p")?;
    }

    Ok(())
}

fn write_textvec(region: &TextVec,
                 xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    if let Some(region) = region.vec() {
        for c in region {
            match c {
                TextElem::Start(ref t) => {
                    xml_out.elem(&t.tag())?;
                    if let Some(attr) = t.attr_map() {
                        for (k, v) in attr.iter() {
                            xml_out.attr_esc(k.as_ref(), v.as_ref())?;
                        }
                    }
                }
                TextElem::Empty(t) => {
                    xml_out.empty(t.tag().as_str())?;
                    if let Some(attr) = t.attr_map() {
                        for (k, v) in attr.iter() {
                            xml_out.attr_esc(k.as_ref(), v.as_ref())?;
                        }
                    }
                }
                TextElem::Text(t) => {
                    xml_out.text_esc(t.as_str())?;
                }
                TextElem::End(t) => {
                    xml_out.end_elem(t.as_str())?;
                }
            }
        }
    }

    Ok(())
}


