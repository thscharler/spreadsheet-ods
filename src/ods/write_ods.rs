use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::{File, rename};
use std::io;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use string_cache::DefaultAtom;

use chrono::NaiveDateTime;
use quick_xml::events::{BytesDecl, Event};
use zip::write::FileOptions;

use crate::{Family, Origin, FormatType, SCell, Sheet, Style, Value, ValueFormat, ValueType, WorkBook, FontDecl};
use crate::ods::error::OdsError;
use crate::ods::temp_zip::TempZip;
use crate::ods::xml_util;

// this did not work out as expected ...
// TODO: find out why this breaks content.xml
// type OdsWriter = zip::ZipWriter<BufWriter<File>>;
// type XmlOdsWriter<'a> = quick_xml::Writer<&'a mut zip::ZipWriter<BufWriter<File>>>;

type OdsWriter = TempZip;
type XmlOdsWriter<'a> = quick_xml::Writer<&'a mut OdsWriter>;

/// Writes the ODS file.
pub fn write_ods<P: AsRef<Path>>(book: &WorkBook, ods_path: P) -> Result<(), OdsError> {
    write_ods_clean(book, ods_path, true)?;
    Ok(())
}

/// Writes the ODS file. The parameter clean indicates the cleanup of the
/// temp files at the end.
pub fn write_ods_clean<P: AsRef<Path>>(book: &WorkBook, ods_path: P, clean: bool) -> Result<(), OdsError> {
    let orig_bak = if let Some(ods_orig) = &book.file {
        let mut orig_bak = ods_orig.clone();
        orig_bak.set_extension("bak");
        rename(&ods_orig, &orig_bak)?;
        Some(orig_bak)
    } else {
        None
    };

    // let zip_file = File::create(ods_path)?;
    // let mut zip_writer = zip::ZipWriter::new(io::BufWriter::new(zip_file));
    let mut zip_writer = TempZip::new(ods_path.as_ref());

    let mut file_set = HashSet::<String>::new();
    //
    if let Some(orig_bak) = orig_bak {
        copy_workbook(&orig_bak, &mut file_set, &mut zip_writer)?;
    }

    write_mimetype(&mut zip_writer, &mut file_set)?;
    write_manifest(&mut zip_writer, &mut file_set)?;
    write_manifest_rdf(&mut zip_writer, &mut file_set)?;
    write_meta(&mut zip_writer, &mut file_set)?;
    //write_settings(&mut zip_writer, &mut file_set)?;
    //write_configurations(&mut zip_writer, &mut file_set)?;

    write_ods_styles(&mut zip_writer, &mut file_set)?;
    write_ods_content(&book, &mut zip_writer, &mut file_set)?;

    zip_writer.zip()?;
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
            zip_writer.start_file(zip_entry.name(), FileOptions::default())?;
            let mut buf: [u8; 1024] = [0; 1024];
            loop {
                let n = zip_entry.read(&mut buf)?;
                if n == 0 {
                    break;
                } else {
                    zip_writer.write_all(&buf[0..n])?;
                }
            }
        }
    }

    Ok(())
}

fn write_mimetype(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), io::Error> {
    if !file_set.contains("mimetype") {
        file_set.insert(String::from("mimetype"));

        zip_out.start_file("mimetype", FileOptions::default().compression_method(zip::CompressionMethod::Stored))?;

        let mime = "application/vnd.oasis.opendocument.spreadsheet";
        zip_out.write_all(mime.as_bytes())?;
    }

    Ok(())
}

fn write_manifest(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("META-INF/manifest.xml") {
        file_set.insert(String::from("META-INF/manifest.xml"));

        zip_out.add_directory("META-INF", FileOptions::default())?;
        zip_out.start_file("META-INF/manifest.xml", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new_with_indent(zip_out, 32, 1);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;

        xml_out.write_event(xml_util::start_a("manifest:manifest", vec![
            ("xmlns:manifest", String::from("urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")),
            ("manifest:version", String::from("1.2")),
        ]))?;

        xml_out.write_event(xml_util::empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("/")),
            ("manifest:version", String::from("1.2")),
            ("manifest:media-type", String::from("application/vnd.oasis.opendocument.spreadsheet")),
        ]))?;
//        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
//            ("manifest:full-path", String::from("Configurations2/")),
//            ("manifest:media-type", String::from("application/vnd.sun.xml.ui.configuration")),
//        ]))?;
        xml_out.write_event(xml_util::empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("manifest.rdf")),
            ("manifest:media-type", String::from("application/rdf+xml")),
        ]))?;
        xml_out.write_event(xml_util::empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("styles.xml")),
            ("manifest:media-type", String::from("text/xml")),
        ]))?;
        xml_out.write_event(xml_util::empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("meta.xml")),
            ("manifest:media-type", String::from("text/xml")),
        ]))?;
        xml_out.write_event(xml_util::empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("content.xml")),
            ("manifest:media-type", String::from("text/xml")),
        ]))?;
//        xml_out.write_event(xml::xml_empty_a("manifest:file-entry", vec![
//            ("manifest:full-path", String::from("settings.xml")),
//            ("manifest:media-type", String::from("text/xml")),
//        ]))?;
        xml_out.write_event(xml_util::end("manifest:manifest"))?;
    }

    Ok(())
}

fn write_manifest_rdf(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("manifest.rdf") {
        file_set.insert(String::from("manifest.rdf"));

        zip_out.start_file("manifest.rdf", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new(zip_out);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
        xml_out.write(b"\n")?;

        xml_out.write_event(xml_util::start_a("rdf:RDF", vec![
            ("xmlns:rdf", String::from("http://www.w3.org/1999/02/22-rdf-syntax-ns#")),
        ]))?;

        xml_out.write_event(xml_util::start_a("rdf:Description", vec![
            ("rdf:about", String::from("content.xml")),
        ]))?;
        xml_out.write_event(xml_util::empty_a("rdf:type", vec![
            ("rdf:resource", String::from("http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile")),
        ]))?;
        xml_out.write_event(xml_util::end("rdf:Description"))?;

        xml_out.write_event(xml_util::start_a("rdf:Description", vec![
            ("rdf:about", String::from("")),
        ]))?;
        xml_out.write_event(xml_util::empty_a("ns0:hasPart", vec![
            ("xmlns:ns0", String::from("http://docs.oasis-open.org/ns/office/1.2/meta/pkg#")),
            ("rdf:resource", String::from("content.xml")),
        ]))?;
        xml_out.write_event(xml_util::end("rdf:Description"))?;

        xml_out.write_event(xml_util::start_a("rdf:Description", vec![
            ("rdf:about", String::from("")),
        ]))?;
        xml_out.write_event(xml_util::empty_a("rdf:type", vec![
            ("rdf:resource", String::from("http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document")),
        ]))?;
        xml_out.write_event(xml_util::end("rdf:Description"))?;

        xml_out.write_event(xml_util::end("rdf:RDF"))?;
    }

    Ok(())
}

fn write_meta(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("meta.xml") {
        file_set.insert(String::from("meta.xml"));

        zip_out.start_file("meta.xml", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new(zip_out);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
        xml_out.write(b"\n")?;

        xml_out.write_event(xml_util::start_a("office:document-meta", vec![
            ("xmlns:meta", String::from("urn:oasis:names:tc:opendocument:xmlns:meta:1.0")),
            ("xmlns:office", String::from("urn:oasis:names:tc:opendocument:xmlns:office:1.0")),
            ("office:version", String::from("1.2")),
        ]))?;

        xml_out.write_event(xml_util::start("office:meta"))?;

        xml_out.write_event(xml_util::start("meta:generator"))?;
        xml_out.write_event(xml_util::text("spreadsheet-ods 0.1.0"))?;
        xml_out.write_event(xml_util::end("meta:generator"))?;

        xml_out.write_event(xml_util::start("meta:creation-date"))?;
        let s = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?;
        let d = NaiveDateTime::from_timestamp(s.as_secs() as i64, 0);
        xml_out.write_event(xml_util::text(&d.format("%Y-%m-%dT%H:%M:%S%.f").to_string()))?;
        xml_out.write_event(xml_util::end("meta:creation-date"))?;

        xml_out.write_event(xml_util::start("meta:editing-duration"))?;
        xml_out.write_event(xml_util::text("P0D"))?;
        xml_out.write_event(xml_util::end("meta:editing-duration"))?;

        xml_out.write_event(xml_util::start("meta:editing-cycles"))?;
        xml_out.write_event(xml_util::text("1"))?;
        xml_out.write_event(xml_util::end("meta:editing-cycles"))?;

        xml_out.write_event(xml_util::start("meta:initial-creator"))?;
        xml_out.write_event(xml_util::text(&username::get_user_name().unwrap()))?;
        xml_out.write_event(xml_util::end("meta:initial-creator"))?;

        xml_out.write_event(xml_util::end("office:meta"))?;

        xml_out.write_event(xml_util::end("office:document-meta"))?;
    }

    Ok(())
}

//fn write_settings(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
//    if !file_set.contains("settings.xml") {
//        file_set.insert(String::from("settings.xml"));
//
//        zip_out.start_file("settings.xml", FileOptions::default())?;
//
//        let mut xml_out = quick_xml::Writer::new(zip_out);
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

fn write_ods_styles(zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("styles.xml") {
        file_set.insert(String::from("styles.xml"));

        zip_out.start_file("styles.xml", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new(zip_out);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
        xml_out.write(b"\n")?;

        xml_out.write_event(xml_util::start_a("office:document-styles", vec![
            ("xmlns:meta", String::from("urn:oasis:names:tc:opendocument:xmlns:meta:1.0")),
            ("xmlns:office", String::from("urn:oasis:names:tc:opendocument:xmlns:office:1.0")),
            ("xmlns:fo", String::from("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")),
            ("xmlns:style", String::from("urn:oasis:names:tc:opendocument:xmlns:style:1.0")),
            ("xmlns:text", String::from("urn:oasis:names:tc:opendocument:xmlns:text:1.0")),
            ("xmlns:dr3d", String::from("urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")),
            ("xmlns:svg", String::from("urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")),
            ("xmlns:chart", String::from("urn:oasis:names:tc:opendocument:xmlns:chart:1.0")),
            ("xmlns:table", String::from("urn:oasis:names:tc:opendocument:xmlns:table:1.0")),
            ("xmlns:number", String::from("urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")),
            ("xmlns:of", String::from("urn:oasis:names:tc:opendocument:xmlns:of:1.2")),
            ("xmlns:calcext", String::from("urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0")),
            ("xmlns:loext", String::from("urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")),
            ("xmlns:field", String::from("urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")),
            ("xmlns:form", String::from("urn:oasis:names:tc:opendocument:xmlns:form:1.0")),
            ("xmlns:script", String::from("urn:oasis:names:tc:opendocument:xmlns:script:1.0")),
            ("xmlns:presentation", String::from("urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")),
            ("office:version", String::from("1.2")),
        ]))?;

        // TODO: read and write global styles

        xml_out.write_event(xml_util::end("office:document-styles"))?;
    }

    Ok(())
}

fn write_ods_content(book: &WorkBook, zip_out: &mut OdsWriter, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    file_set.insert(String::from("content.xml"));

    zip_out.start_file("content.xml", FileOptions::default())?;

    let mut xml_out = quick_xml::Writer::new_with_indent(zip_out, b' ', 1);

    xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
    xml_out.write(b"\n")?;
    xml_out.write_event(xml_util::start_a("office:document-content", vec![
        ("xmlns:presentation", String::from("urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")),
        ("xmlns:grddl", String::from("http://www.w3.org/2003/g/data-view#")),
        ("xmlns:xhtml", String::from("http://www.w3.org/1999/xhtml")),
        ("xmlns:xsi", String::from("http://www.w3.org/2001/XMLSchema-instance")),
        ("xmlns:xsd", String::from("http://www.w3.org/2001/XMLSchema")),
        ("xmlns:xforms", String::from("http://www.w3.org/2002/xforms")),
        ("xmlns:dom", String::from("http://www.w3.org/2001/xml-events")),
        ("xmlns:script", String::from("urn:oasis:names:tc:opendocument:xmlns:script:1.0")),
        ("xmlns:form", String::from("urn:oasis:names:tc:opendocument:xmlns:form:1.0")),
        ("xmlns:math", String::from("http://www.w3.org/1998/Math/MathML")),
        ("xmlns:draw", String::from("urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")),
        ("xmlns:dr3d", String::from("urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")),
        ("xmlns:text", String::from("urn:oasis:names:tc:opendocument:xmlns:text:1.0")),
        ("xmlns:style", String::from("urn:oasis:names:tc:opendocument:xmlns:style:1.0")),
        ("xmlns:meta", String::from("urn:oasis:names:tc:opendocument:xmlns:meta:1.0")),
        ("xmlns:ooo", String::from("http://openoffice.org/2004/office")),
        ("xmlns:loext", String::from("urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")),
        ("xmlns:svg", String::from("urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")),
        ("xmlns:of", String::from("urn:oasis:names:tc:opendocument:xmlns:of:1.2")),
        ("xmlns:office", String::from("urn:oasis:names:tc:opendocument:xmlns:office:1.0")),
        ("xmlns:fo", String::from("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")),
        ("xmlns:field", String::from("urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")),
        ("xmlns:xlink", String::from("http://www.w3.org/1999/xlink")),
        ("xmlns:formx", String::from("urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0")),
        ("xmlns:dc", String::from("http://purl.org/dc/elements/1.1/")),
        ("xmlns:chart", String::from("urn:oasis:names:tc:opendocument:xmlns:chart:1.0")),
        ("xmlns:rpt", String::from("http://openoffice.org/2005/report")),
        ("xmlns:table", String::from("urn:oasis:names:tc:opendocument:xmlns:table:1.0")),
        ("xmlns:css3t", String::from("http://www.w3.org/TR/css3-text/")),
        ("xmlns:number", String::from("urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")),
        ("xmlns:ooow", String::from("http://openoffice.org/2004/writer")),
        ("xmlns:oooc", String::from("http://openoffice.org/2004/calc")),
        ("xmlns:tableooo", String::from("http://openoffice.org/2009/table")),
        ("xmlns:calcext", String::from("urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0")),
        ("xmlns:drawooo", String::from("http://openoffice.org/2010/draw")),
        ("office:version", String::from("1.2")),
    ]))?;
    xml_out.write_event(xml_util::empty("office:scripts"))?;

    xml_out.write_event(xml_util::start("office:font-face-decls"))?;
    write_font_decl(&book.fonts, Origin::Content, &mut xml_out)?;
    xml_out.write_event(xml_util::end("office:font-face-decls"))?;

    xml_out.write_event(xml_util::start("office:automatic-styles"))?;
    write_styles(&book.styles, Origin::Content, &mut xml_out)?;
    write_value_styles(&book.formats, Origin::Content, &mut xml_out)?;
    xml_out.write_event(xml_util::end("office:automatic-styles"))?;

    xml_out.write_event(xml_util::start("office:body"))?;
    xml_out.write_event(xml_util::start("office:spreadsheet"))?;

    for sheet in &book.sheets {
        write_sheet(&book, &sheet, &mut xml_out)?;
    }

    xml_out.write_event(xml_util::end("office:spreadsheet"))?;
    xml_out.write_event(xml_util::end("office:body"))?;
    xml_out.write_event(xml_util::end("office:document-content"))?;

    Ok(())
}

fn write_sheet(book: &WorkBook, sheet: &Sheet, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    attr.push(("table:name", sheet.name.to_string()));
    if let Some(style) = &sheet.style {
        attr.push(("table:style-name", style.to_string()));
    }
    xml_out.write_event(xml_util::start_a("table:table", attr))?;

    let max_cell = sheet.used_grid_size();

    write_table_columns(&sheet, &max_cell, xml_out)?;

    // table-row + table-cell
    let mut first_cell = true;
    let mut last_r: usize = 0;
    let mut last_c: usize = 0;

    for ((cur_row, cur_col), cell) in sheet.data.iter() {

        // There may be a lot of gaps of any kind in our data.
        // In the XML format there is no cell identification, every gap
        // must be filled with empty rows/columns. For this we need some
        // calculations.

        // For the repeat-counter we need to look forward.
        // Works nicely with the range operator :-)
        let (next_r, next_c) =
            if let Some(((next_r, next_c), _)) = sheet.data.range((*cur_row, cur_col + 1)..).next() {
                (*next_r, *next_c)
            } else {
                (max_cell.0, max_cell.1)
            };

        // Looking forward row-wise.
        let forward_dr = next_r as i32 - *cur_row as i32;

        // Column deltas are only relevant in the same row. Except we need to
        // fill up to max used columns.
        let forward_dc = if forward_dr >= 1 {
            max_cell.1 as i32 - *cur_col as i32
        } else {
            next_c as i32 - *cur_col as i32
        };

        // Looking backward row-wise.
        let backward_dr = *cur_row as i32 - last_r as i32;
        // When a row changes our delta is from zero to cur_col.
        let backward_dc = if backward_dr >= 1 {
            *cur_col as i32
        } else {
            *cur_col as i32 - last_c as i32
        };

        // After the first cell there is always an open row tag.
        if backward_dr > 0 && !first_cell {
            xml_out.write_event(xml_util::end("table:table-row"))?;
        }

        // Any empty rows before this one?
        if backward_dr > 1 {
            write_empty_rows_before(first_cell, backward_dr, &max_cell, xml_out)?;
        }

        // Start a new row if there is a delta or we are at the start
        if backward_dr > 0 || first_cell {
            write_start_current_row(sheet, *cur_row, backward_dc, xml_out)?;
        }

        // And now to something completely different ...
        write_cell(book, cell, xml_out)?;

        // There may be some blank cells until the next one, but only one less the forward.
        if forward_dc > 1 {
            write_empty_cells(forward_dc, xml_out)?;
        }

        first_cell = false;
        last_r = *cur_row;
        last_c = *cur_col;
    }

    if !first_cell {
        xml_out.write_event(xml_util::end("table:table-row"))?;
    }

    xml_out.write_event(xml_util::end("table:table"))?;

    Ok(())
}

fn write_empty_cells(forward_dc: i32, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    attr.push(("table:number-columns-repeated", (forward_dc - 1).to_string()));
    xml_out.write_event(xml_util::empty_a("table:table-cell", attr))?;

    Ok(())
}

fn write_start_current_row(sheet: &Sheet,
                           cur_row: usize,
                           backward_dc: i32,
                           xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    if let Some(row) = sheet.rows.get(&cur_row) {
        if let Some(style) = &row.style {
            attr.push(("table:style-name", style.to_string()));
        }
    }

    xml_out.write_event(xml_util::start_a("table:table-row", attr))?;

    // Might not be the first column in this row.
    if backward_dc > 0 {
        xml_out.write_event(xml_util::empty_a("table:table-cell", vec![
            ("table:number-columns-repeated", backward_dc.to_string()),
        ]))?;
    }

    Ok(())
}

fn write_empty_rows_before(first_cell: bool,
                           backward_dr: i32,
                           max_cell: &(usize, usize),
                           xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    // Empty rows in between are 1 less than the delta, except at the very start.
    let mut attr = Vec::new();

    let empty_count = if first_cell { backward_dr } else { backward_dr - 1 };
    if empty_count > 1 {
        attr.push(("table:number-rows-repeated", empty_count.to_string()));
    }

    xml_out.write_event(xml_util::start_a("table:table-row", attr))?;

    // We fill the empty spaces completely up to max columns.
    xml_out.write_event(xml_util::empty_a("table:table-cell", vec![
        ("table:number-columns-repeated", max_cell.1.to_string()),
    ]))?;

    xml_out.write_event(xml_util::end("table:table-row"))?;

    Ok(())
}

fn write_table_columns(sheet: &Sheet,
                       max_cell: &(usize, usize),
                       xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    // table:table-column
    for c in sheet.columns.keys() {
        // For the repeat-counter we need to look forward.
        // Works nicely with the range operator :-)
        let next_c = if let Some((next_c, _)) = sheet.columns.range(c + 1..).next() {
            *next_c
        } else {
            max_cell.1
        };

        let dc = next_c as i32 - *c as i32;

        let column = &sheet.columns[c];

        let mut attr = Vec::new();
        if dc > 1 {
            attr.push(("table:number-columns-repeated", dc.to_string()));
        }
        if let Some(style) = &column.style {
            attr.push(("table:style-name", style.to_string()));
        }
        if let Some(style) = &column.def_cell_style {
            attr.push(("table:default-cell-style-name", style.to_string()));
        }
        xml_out.write_event(xml_util::empty_a("table:table-column", attr))?;
    }

    Ok(())
}

fn write_cell(book: &WorkBook,
              cell: &SCell,
              xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    let mut content: String = String::from("");
    if let Some(formula) = &cell.formula {
        attr.push(("table:formula", formula.to_string()));
    }

    if let Some(style) = &cell.style {
        attr.push(("table:style-name", style.to_string()));
    } else if let Some(value) = &cell.value {
        if let Some(style) = book.def_style(value.value_type()) {
            attr.push(("table:style-name", style.to_string()));
        }
    }

    // Might not yield a useful result. Could not exist, or be in styles.xml
    // which I don't read. Defaulting to to_string() seems reasonable.
    let value_style = if let Some(style_name) = &cell.style {
        book.find_value_style(style_name)
    } else {
        None
    };

    let mut is_empty = false;
    match &cell.value {
        None => {
            is_empty = true;
        }
        Some(Value::Text(s)) => {
            attr.push(("office:value-type", String::from("string")));
            if let Some(value_style) = value_style {
                content = value_style.format_str(s);
            } else {
                content = s.to_string();
            }
        }
        Some(Value::DateTime(d)) => {
            attr.push(("office:value-type", String::from("date")));
            attr.push(("office:date-value", d.format("%Y-%m-%dT%H:%M:%S%.f").to_string()));
            if let Some(value_style) = value_style {
                content = value_style.format_datetime(d);
            } else {
                content = d.format("%d.%m.%Y").to_string();
            }
        }
        Some(Value::TimeDuration(d)) => {
            attr.push(("office:value-type", String::from("time")));

            let mut buf = String::from("PT");
            buf.push_str(&d.num_hours().to_string());
            buf.push_str("H");
            buf.push_str(&(d.num_minutes() % 60).to_string());
            buf.push_str("M");
            buf.push_str(&(d.num_seconds() % 60).to_string());
            buf.push_str(".");
            buf.push_str(&(d.num_milliseconds() % 1000).to_string());
            buf.push_str("S");

            attr.push(("office:time-value", buf));
            if let Some(value_style) = value_style {
                content = value_style.format_time_duration(d);
            } else {
                content.push_str(&d.num_hours().to_string());
                content.push_str(":");
                content.push_str(&(d.num_minutes() % 60).to_string());
                content.push_str(":");
                content.push_str(&(d.num_seconds() % 60).to_string());
                content.push_str(".");
                content.push_str(&(d.num_milliseconds() % 1000).to_string());
            }
        }
        Some(Value::Boolean(b)) => {
            attr.push(("office:value-type", String::from("boolean")));
            attr.push(("office:boolean-value", b.to_string()));
            if let Some(value_style) = value_style {
                content = value_style.format_boolean(*b);
            } else {
                content = b.to_string();
            }
        }
        Some(Value::Currency(c, v)) => {
            attr.push(("office:value-type", String::from("currency")));
            attr.push(("office:currency", c.to_string()));
            attr.push(("office:value", v.to_string()));
            if let Some(value_style) = value_style {
                content = value_style.format_float(*v);
            } else {
                content.push_str(c);
                content.push_str(" ");
                content.push_str(&v.to_string());
            }
        }
        Some(Value::Number(v)) => {
            attr.push(("office:value-type", String::from("float")));
            attr.push(("office:value", v.to_string()));
            if let Some(value_style) = value_style {
                content = value_style.format_float(*v);
            } else {
                content = v.to_string();
            }
        }
        Some(Value::Percentage(v)) => {
            attr.push(("office:value-type", String::from("percentage")));
            attr.push(("office:value", format!("{}%", v)));
            if let Some(value_style) = value_style {
                content = value_style.format_float(*v * 100.0);
            } else {
                content = (v * 100.0).to_string();
            }
        }
    }

    if !is_empty {
        xml_out.write_event(xml_util::start_a("table:table-cell", attr))?;
        xml_out.write_event(xml_util::start("text:p"))?;
        xml_out.write_event(xml_util::text(&content))?;
        xml_out.write_event(xml_util::end("text:p"))?;
        xml_out.write_event(xml_util::end("table:table-cell"))?;
    } else {
        xml_out.write_event(xml_util::empty_a("table:table-cell", attr))?;
    }

    Ok(())
}

fn write_font_decl(fonts: &BTreeMap<String, FontDecl>, origin: Origin, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for font in fonts.values().filter(|s| s.origin == origin) {
        let mut attr: HashMap<DefaultAtom, String> = HashMap::new();

        attr.insert(DefaultAtom::from("style:name"), font.name.to_string());

        if let Some(prp) = &font.prp {
            for (k, v) in prp {
                attr.insert(k.clone(), v.to_string());
            }
        }

        xml_out.write_event(xml_util::start_m("style:style", &attr))?;
    }
    Ok(())
}

fn write_styles(styles: &BTreeMap<String, Style>, origin: Origin, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for style in styles.values().filter(|s| s.origin == origin) {
        let mut attr: Vec<(&str, String)> = Vec::new();

        attr.push(("style:name", style.name.to_string()));
        let family = match style.family {
            Family::Table => "table",
            Family::TableColumn => "table-column",
            Family::TableRow => "table-row",
            Family::TableCell => "table-cell",
            Family::None => "",
        };
        attr.push(("style:family", family.to_string()));
        if let Some(display_name) = &style.display_name {
            attr.push(("style:display-name", display_name.to_string()));
        }
        if let Some(parent) = &style.parent {
            attr.push(("style:parent-style-name", parent.to_string()));
        }
        if let Some(value_style) = &style.value_style {
            attr.push(("style:data-style-name", value_style.to_string()));
        }
        xml_out.write_event(xml_util::start_a("style:style", attr))?;

        if let Some(table_cell_prp) = &style.table_cell_prp {
            xml_out.write_event(xml_util::empty_m("style:table-cell-properties", &table_cell_prp))?;
        }
        if let Some(table_col_prp) = &style.table_col_prp {
            xml_out.write_event(xml_util::empty_m("style:table-column-properties", &table_col_prp))?;
        }
        if let Some(table_row_prp) = &style.table_row_prp {
            xml_out.write_event(xml_util::empty_m("style:table-row-properties", &table_row_prp))?;
        }
        if let Some(table_prp) = &style.table_prp {
            xml_out.write_event(xml_util::empty_m("style:table-properties", &table_prp))?;
        }
        if let Some(paragraph_prp) = &style.paragraph_prp {
            xml_out.write_event(xml_util::empty_m("style:paragraph-properties", &paragraph_prp))?;
        }
        if let Some(text_prp) = &style.text_prp {
            xml_out.write_event(xml_util::empty_m("style:text-properties", &text_prp))?;
        }

        xml_out.write_event(xml_util::end("style:style"))?;
    }

    Ok(())
}

fn write_value_styles(styles: &BTreeMap<String, ValueFormat>, origin: Origin, xml_out: &mut XmlOdsWriter) -> Result<(), OdsError> {
    for style in styles.values().filter(|s| s.origin == origin) {
        let tag = match style.v_type {
            ValueType::Boolean => "number:boolean-style",
            ValueType::Number => "number:number-style",
            ValueType::Text => "number:text-style",
            ValueType::TimeDuration => "number:time-style",
            ValueType::Percentage => "number:percentage-style",
            ValueType::Currency => "number:currency-style",
            ValueType::DateTime => "number:date-style",
        };

        let mut attr: HashMap<DefaultAtom, String> = HashMap::new();
        attr.insert(DefaultAtom::from("style:name"), style.name.to_string());
        if let Some(m) = &style.prp {
            for (k, v) in m {
                attr.insert(k.clone(), v.to_string());
            }
        }
        xml_out.write_event(xml_util::start_m(tag, &attr))?;

        if let Some(parts) = style.parts() {
            for part in parts {
                let part_tag = match part.ftype {
                    FormatType::Boolean => "number:boolean",
                    FormatType::Number => "number:number",
                    FormatType::Scientific => "number:scientific-number",
                    FormatType::CurrencySymbol => "number:currency-symbol",
                    FormatType::Day => "number:day",
                    FormatType::Month => "number:month",
                    FormatType::Year => "number:year",
                    FormatType::Era => "number:era",
                    FormatType::DayOfWeek => "number:day-of-week",
                    FormatType::WeekOfYear => "number:week-of-year",
                    FormatType::Quarter => "number:quarter",
                    FormatType::Hours => "number:hours",
                    FormatType::Minutes => "number:minutes",
                    FormatType::Seconds => "number:seconds",
                    FormatType::Fraction => "number:fraction",
                    FormatType::AmPm => "number:am-pm",
                    FormatType::EmbeddedText => "number:embedded-text",
                    FormatType::Text => "number:text",
                    FormatType::TextContent => "number:text-content",
                    FormatType::StyleText => "style:text",
                    FormatType::StyleMap => "style:map",
                };

                if part.ftype == FormatType::Text || part.ftype == FormatType::CurrencySymbol {
                    xml_out.write_event(xml_util::start_o(part_tag, part.prp.as_ref()))?;
                    if let Some(content) = &part.content {
                        xml_out.write_event(xml_util::text(content))?;
                    }
                    xml_out.write_event(xml_util::end(part_tag))?;
                } else {
                    xml_out.write_event(xml_util::empty_o(part_tag, part.prp.as_ref()))?;
                }
            }
        }
        xml_out.write_event(xml_util::end(tag))?;
    }

    Ok(())
}