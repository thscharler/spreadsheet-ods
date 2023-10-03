use std::borrow::Cow;
use std::fs::File;
use std::io;
use std::io::{BufWriter, Cursor, Seek, Write};
use std::path::Path;

use zip::write::FileOptions;

use crate::config::{ConfigItem, ConfigItemType, ConfigValue};
use crate::error::OdsError;
use crate::format::FormatPartType;
use crate::io::format::{format_duration2, format_validation_condition};
use crate::io::xmlwriter::XmlWriter;
use crate::io::zip_out::ZipOut;
use crate::manifest::Manifest;
use crate::metadata::MetaValue;
use crate::refs::{format_cellranges, CellRange};
use crate::style::{
    CellStyle, ColStyle, FontFaceDecl, GraphicStyle, HeaderFooter, MasterPage, PageStyle,
    ParagraphStyle, RowStyle, RubyStyle, StyleOrigin, StyleUse, TableStyle, TextStyle,
};
use crate::validation::ValidationDisplay;
use crate::xmltree::{XmlContent, XmlTag};
use crate::{
    CellContentRef, EventListener, Length, Script, Sheet, Value, ValueFormatTrait, ValueType,
    Visibility, WorkBook,
};
use crate::{HashMap, NamespaceMap};

type OdsXmlWriter<'a> = XmlWriter<BufWriter<Box<dyn Write + 'a>>>;

const DATETIME_FORMAT: &str = "%Y-%m-%dT%H:%M:%S%.f";

/// Writes the ODS file into a supplied buffer.
pub fn write_ods_buf_uncompressed(book: &mut WorkBook, buf: Vec<u8>) -> Result<Vec<u8>, OdsError> {
    let zip_writer = ZipOut::<Cursor<Vec<u8>>>::new_buf_uncompressed(buf)?;
    Ok(write_ods_impl(book, zip_writer)?.into_inner())
}

/// Writes the ODS file into a supplied buffer.
pub fn write_ods_buf(book: &mut WorkBook, buf: Vec<u8>) -> Result<Vec<u8>, OdsError> {
    let zip_writer = ZipOut::<Cursor<Vec<u8>>>::new_buf(buf)?;
    Ok(write_ods_impl(book, zip_writer)?.into_inner())
}

/// Writes the ODS file to the given Write.
pub fn write_ods_to<T: Write + Seek>(book: &mut WorkBook, ods: T) -> Result<(), OdsError> {
    let zip_writer = ZipOut::new_to(ods)?;
    write_ods_impl(book, zip_writer)?;
    Ok(())
}

/// Writes the ODS file.
pub fn write_ods<P: AsRef<Path>>(book: &mut WorkBook, ods_path: P) -> Result<(), OdsError> {
    let zip_writer = ZipOut::<File>::new_file(ods_path.as_ref())?;
    write_ods_impl(book, zip_writer)?;
    Ok(())
}

/// Writes the FODS file into a supplied buffer.
pub fn write_fods_buf(book: &mut WorkBook, mut write: Vec<u8>) -> Result<Vec<u8>, OdsError> {
    write_fods_impl(book, Box::new(&mut write))?;
    Ok(write)
}

/// Writes the FODS file to the given Write.
pub fn write_fods_to<T: Write + Seek>(book: &mut WorkBook, write: T) -> Result<(), OdsError> {
    write_fods_impl(book, Box::new(write))?;
    Ok(())
}

/// Writes the FODS file.
pub fn write_fods<P: AsRef<Path>>(book: &mut WorkBook, ods_path: P) -> Result<(), OdsError> {
    let write = File::create(ods_path)?;
    write_fods_impl(book, Box::new(write))?;
    Ok(())
}

/// Writes the ODS file.
///
/// All the parts are written to a temp directory and then zipped together.
///
fn write_fods_impl(book: &mut WorkBook, writer: Box<dyn Write + '_>) -> Result<(), OdsError> {
    sanity_checks(book)?;
    sync(book)?;

    convert(book)?;

    let mut xml_out = XmlWriter::new(BufWriter::new(writer)).line_break(true);
    write_fods_content(book, &mut xml_out)?;

    Ok(())
}

fn convert(book: &mut WorkBook) -> Result<(), OdsError> {
    for v in book.tablestyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.rowstyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.colstyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.cellstyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.paragraphstyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.textstyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.rubystyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.graphicstyles.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }

    for v in book.formats_boolean.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.formats_number.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.formats_percentage.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.formats_currency.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.formats_text.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.formats_datetime.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }
    for v in book.formats_timeduration.values_mut() {
        v.set_origin(StyleOrigin::Content);
    }

    Ok(())
}

fn write_fods_content(book: &mut WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    let xmlns = book
        .xmlns
        .entry("meta.xml".into())
        .or_insert_with(NamespaceMap::new);

    xmlns.insert_str(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    );
    xmlns.insert_str("xmlns:ooo", "http://openoffice.org/2004/office");
    xmlns.insert_str(
        "xmlns:fo",
        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    );
    xmlns.insert_str("xmlns:xlink", "http://www.w3.org/1999/xlink");
    xmlns.insert_str(
        "xmlns:config",
        "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    );
    xmlns.insert_str("xmlns:dc", "http://purl.org/dc/elements/1.1/");
    xmlns.insert_str(
        "xmlns:meta",
        "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    );
    xmlns.insert_str(
        "xmlns:style",
        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    );
    xmlns.insert_str(
        "xmlns:text",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    );
    xmlns.insert_str("xmlns:rpt", "http://openoffice.org/2005/report");
    xmlns.insert_str(
        "xmlns:draw",
        "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    );
    xmlns.insert_str(
        "xmlns:dr3d",
        "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    );
    xmlns.insert_str(
        "xmlns:svg",
        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    );
    xmlns.insert_str(
        "xmlns:chart",
        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    );
    xmlns.insert_str(
        "xmlns:table",
        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    );
    xmlns.insert_str(
        "xmlns:number",
        "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    );
    xmlns.insert_str("xmlns:ooow", "http://openoffice.org/2004/writer");
    xmlns.insert_str("xmlns:oooc", "http://openoffice.org/2004/calc");
    xmlns.insert_str("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");
    xmlns.insert_str("xmlns:xforms", "http://www.w3.org/2002/xforms");
    xmlns.insert_str("xmlns:tableooo", "http://openoffice.org/2009/table");
    xmlns.insert_str(
        "xmlns:calcext",
        "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    );
    xmlns.insert_str("xmlns:drawooo", "http://openoffice.org/2010/draw");
    xmlns.insert_str(
        "xmlns:loext",
        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    );
    xmlns.insert_str(
        "xmlns:field",
        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
    );
    xmlns.insert_str("xmlns:math", "http://www.w3.org/1998/Math/MathML");
    xmlns.insert_str(
        "xmlns:form",
        "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    );
    xmlns.insert_str(
        "xmlns:script",
        "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    );
    xmlns.insert_str(
        "xmlns:formx",
        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
    );
    xmlns.insert_str("xmlns:dom", "http://www.w3.org/2001/xml-events");
    xmlns.insert_str("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
    xmlns.insert_str("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
    xmlns.insert_str("xmlns:xhtml", "http://www.w3.org/1999/xhtml");
    xmlns.insert_str("xmlns:grddl", "http://www.w3.org/2003/g/data-view#");
    xmlns.insert_str("xmlns:css3t", "http://www.w3.org/TR/css3-text/");
    xmlns.insert_str(
        "xmlns:presentation",
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    );

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document")?;
    write_xmlns(xmlns, xml_out)?;
    xml_out.attr_esc("office:version", book.version())?;
    xml_out.attr_esc(
        "office:mimetype",
        "application/vnd.oasis.opendocument.spreadsheet",
    )?;

    write_office_meta(book, xml_out)?;
    write_office_settings(book, xml_out)?;
    write_office_scripts(book, xml_out)?;
    write_office_font_face_decls(book, StyleOrigin::Content, xml_out)?;
    write_office_styles(book, StyleOrigin::Content, xml_out)?;
    write_office_automatic_styles(book, StyleOrigin::Content, xml_out)?;
    write_office_master_styles(book, xml_out)?;
    write_office_body(book, xml_out)?;

    xml_out.end_elem("office:document")?;

    xml_out.close()?;

    Ok(())
}

/// Writes the ODS file.
///
/// All the parts are written to a temp directory and then zipped together.
///
fn write_ods_impl<W: Write + Seek>(
    book: &mut WorkBook,
    mut zip_writer: ZipOut<W>,
) -> Result<W, OdsError> {
    sanity_checks(book)?;

    sync(book)?;

    create_manifest(book)?;

    write_ods_mimetype(&mut zip_writer)?;

    zip_writer.add_directory("META-INF", FileOptions::default())?;
    let out = zip_writer.start_file("META-INF/manifest.xml", FileOptions::default())?;
    let write: Box<dyn Write> = Box::new(out);
    write_ods_manifest(book, &mut XmlWriter::new(BufWriter::new(write)))?;

    let out = zip_writer.start_file("meta.xml", FileOptions::default())?;
    let write: Box<dyn Write> = Box::new(out);
    write_ods_metadata(book, &mut XmlWriter::new(BufWriter::new(write)))?;

    let out = zip_writer.start_file("settings.xml", FileOptions::default())?;
    let write: Box<dyn Write> = Box::new(out);
    write_ods_settings(book, &mut XmlWriter::new(BufWriter::new(write)))?;

    let out = zip_writer.start_file("styles.xml", FileOptions::default())?;
    let write: Box<dyn Write> = Box::new(out);
    write_ods_styles(book, &mut XmlWriter::new(BufWriter::new(write)))?;

    let out = zip_writer.start_file("content.xml", FileOptions::default())?;
    let write: Box<dyn Write> = Box::new(out);
    write_ods_content(book, &mut XmlWriter::new(BufWriter::new(write)))?;

    write_ods_extra(book, &mut zip_writer)?;

    Ok(zip_writer.zip()?)
}

/// Sanity checks.
fn sanity_checks(book: &mut WorkBook) -> Result<(), OdsError> {
    if book.sheets.is_empty() {
        return Err(OdsError::Ods("Workbook contains no sheets.".to_string()));
    }
    Ok(())
}

/// - Syncs book.config back to the config tree structure.
/// - Syncs row-heights and col-widths back to the corresponding styles.
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::collapsible_if)]
fn sync(book: &mut WorkBook) -> Result<(), OdsError> {
    // Manifest
    book.metadata.generator = "spreadsheet-ods 0.18.0".to_string();
    book.metadata.document_statistics.table_count = book.sheets.len() as u32;
    let mut cell_count = 0;
    for sheet in book.iter_sheets() {
        cell_count += sheet.data.len() as u32;
    }
    book.metadata.document_statistics.cell_count = cell_count;

    // Config
    let mut config = book.config.detach(0);

    let bc = config.create_path(&[
        ("ooo:view-settings", ConfigItemType::Set),
        ("Views", ConfigItemType::Vec),
        ("0", ConfigItemType::Entry),
    ]);
    if book.config().active_table.is_empty() {
        book.config_mut().active_table = book.sheet(0).name().clone();
    }
    bc.insert("ActiveTable", book.config().active_table.clone());
    bc.insert("HasSheetTabs", book.config().has_sheet_tabs);
    bc.insert("ShowGrid", book.config().show_grid);
    bc.insert("ShowPageBreaks", book.config().show_page_breaks);

    for i in 0..book.num_sheets() {
        let mut sheet = book.detach_sheet(i);

        // Set the column widths.
        for ch in sheet.col_header.values_mut() {
            // Any non default values?
            if ch.width() != Length::Default {
                if ch.style().is_none() {
                    let colstyle = book.add_colstyle(ColStyle::new_empty());
                    ch.set_style(&colstyle);
                }
            }

            // Write back to the style.
            if let Some(style_name) = ch.style() {
                if let Some(style) = book.colstyle_mut(style_name) {
                    if ch.width() == Length::Default {
                        style.set_use_optimal_col_width(true);
                        style.set_col_width(Length::Default);
                    } else {
                        style.set_col_width(ch.width());
                    }
                }
            }
        }

        for rh in sheet.row_header.values_mut() {
            if rh.height() != Length::Default {
                if rh.style().is_none() {
                    let rowstyle = book.add_rowstyle(RowStyle::new_empty());
                    rh.set_style(&rowstyle);
                }
            }

            if let Some(style_name) = rh.style() {
                if let Some(style) = book.rowstyle_mut(style_name) {
                    if rh.height() == Length::Default {
                        style.set_use_optimal_row_height(true);
                        style.set_row_height(Length::Default);
                    } else {
                        style.set_row_height(rh.height());
                    }
                }
            }
        }

        let bc = config.create_path(&[
            ("ooo:view-settings", ConfigItemType::Set),
            ("Views", ConfigItemType::Vec),
            ("0", ConfigItemType::Entry),
            ("Tables", ConfigItemType::Map),
            (sheet.name().as_str(), ConfigItemType::Entry),
        ]);

        bc.insert("CursorPositionX", sheet.config().cursor_x);
        bc.insert("CursorPositionY", sheet.config().cursor_y);
        bc.insert("HorizontalSplitMode", sheet.config().hor_split_mode as i16);
        bc.insert("VerticalSplitMode", sheet.config().vert_split_mode as i16);
        bc.insert("HorizontalSplitPosition", sheet.config().hor_split_pos);
        bc.insert("VerticalSplitPosition", sheet.config().vert_split_pos);
        bc.insert("ActiveSplitRange", sheet.config().active_split_range);
        bc.insert("PositionLeft", sheet.config().position_left);
        bc.insert("PositionRight", sheet.config().position_right);
        bc.insert("PositionTop", sheet.config().position_top);
        bc.insert("PositionBottom", sheet.config().position_bottom);
        bc.insert("ZoomType", sheet.config().zoom_type);
        bc.insert("ZoomValue", sheet.config().zoom_value);
        bc.insert("PageViewZoomValue", sheet.config().page_view_zoom_value);
        bc.insert("ShowGrid", sheet.config().show_grid);

        let bc = config.create_path(&[
            ("ooo:configuration-settings", ConfigItemType::Set),
            ("ScriptConfiguration", ConfigItemType::Map),
            (sheet.name().as_str(), ConfigItemType::Entry),
        ]);
        // maybe this is not accurate. there seem to be cases where the codename
        // is not the same as the table name, but I can't find any uses of the
        // codename anywhere.
        bc.insert("CodeName", sheet.name().as_str().to_string());

        book.attach_sheet(sheet);
    }

    book.config.attach(config);

    Ok(())
}

// Create the standard manifest entries.
fn create_manifest(book: &mut WorkBook) -> Result<(), OdsError> {
    if !book.manifest.contains_key("/") {
        book.add_manifest(Manifest {
            full_path: "/".to_string(),
            version: Some(book.version().clone()),
            media_type: "application/vnd.oasis.opendocument.spreadsheet".to_string(),
            buffer: None,
        });
    }
    if !book.manifest.contains_key("manifest.rdf") {
        book.add_manifest(create_manifest_rdf()?);
    }
    if !book.manifest.contains_key("styles.xml") {
        book.add_manifest(Manifest::new("styles.xml", "text/xml"));
    }
    if !book.manifest.contains_key("meta.xml") {
        book.add_manifest(Manifest::new("meta.xml", "text/xml"));
    }
    if !book.manifest.contains_key("content.xml") {
        book.add_manifest(Manifest::new("content.xml", "text/xml"));
    }
    if !book.manifest.contains_key("settings.xml") {
        book.add_manifest(Manifest::new("settings.xml", "text/xml"));
    }

    Ok(())
}

fn create_manifest_rdf() -> Result<Manifest, OdsError> {
    let mut buf = Vec::new();
    let mut xml_out = XmlWriter::new(&mut buf);

    xml_out.dtd("UTF-8")?;
    xml_out.elem("rdf:RDF")?;
    xml_out.attr_str("xmlns:rdf", "http://www.w3.org/1999/02/22-rdf-syntax-ns#")?;
    xml_out.elem("rdf:Description")?;
    xml_out.attr_str("rdf:about", "content.xml")?;
    xml_out.empty("rdf:type")?;
    xml_out.attr_str(
        "rdf:resource",
        "http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile",
    )?;
    xml_out.end_elem("rdf:Description")?;
    xml_out.elem("rdf:Description")?;
    xml_out.attr_str("rdf:about", "")?;
    xml_out.empty("ns0:hasPart")?;
    xml_out.attr_str(
        "xmlns:ns0",
        "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#",
    )?;
    xml_out.attr_str("rdf:resource", "content.xml")?;
    xml_out.end_elem("rdf:Description")?;
    xml_out.elem("rdf:Description")?;
    xml_out.attr_str("rdf:about", "")?;
    xml_out.empty("rdf:type")?;
    xml_out.attr_str(
        "rdf:resource",
        "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document",
    )?;
    xml_out.end_elem("rdf:Description")?;
    xml_out.end_elem("rdf:RDF")?;
    xml_out.close()?;

    Ok(Manifest::with_buf(
        "manifest.rdf",
        "application/rdf+xml",
        buf,
    ))
}

fn write_ods_mimetype<W: Write + Seek>(zip_writer: &mut ZipOut<W>) -> Result<(), io::Error> {
    let mut out = zip_writer.start_file(
        "mimetype",
        FileOptions::default().compression_method(zip::CompressionMethod::Stored),
    )?;
    out.write_all("application/vnd.oasis.opendocument.spreadsheet".as_bytes())?;
    Ok(())
}

fn write_ods_manifest(book: &WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    xml_out.dtd("UTF-8")?;

    xml_out.elem("manifest:manifest")?;
    xml_out.attr_str(
        "xmlns:manifest",
        "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    )?;
    xml_out.attr_esc("manifest:version", &book.version())?;

    for manifest in book.manifest.values() {
        xml_out.empty("manifest:file-entry")?;
        xml_out.attr_esc("manifest:full-path", &manifest.full_path)?;
        if let Some(version) = &manifest.version {
            xml_out.attr_esc("manifest:version", version)?;
        }
        xml_out.attr_esc("manifest:media-type", &manifest.media_type)?;
    }

    xml_out.end_elem("manifest:manifest")?;

    xml_out.close()?;

    Ok(())
}

fn write_xmlns(xmlns: &NamespaceMap, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    for (k, v) in xmlns.entries() {
        match k {
            Cow::Borrowed(k) => {
                xml_out.attr(k, v.as_ref())?;
            }
            Cow::Owned(k) => {
                xml_out.attr_esc(k, v.as_ref())?;
            }
        }
    }
    Ok(())
}

fn write_ods_metadata(book: &mut WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    let xmlns = book
        .xmlns
        .entry("meta.xml".into())
        .or_insert_with(NamespaceMap::new);

    xmlns.insert_str(
        "xmlns:meta",
        "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    );
    xmlns.insert_str(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    );

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-meta")?;
    write_xmlns(xmlns, xml_out)?;
    xml_out.attr_esc("office:version", book.version())?;

    write_office_meta(book, xml_out)?;

    xml_out.end_elem("office:document-meta")?;

    xml_out.close()?;

    Ok(())
}

fn write_office_meta(book: &WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    xml_out.elem("office:meta")?;

    xml_out.elem_text("meta:generator", &book.metadata.generator)?;
    if !book.metadata.title.is_empty() {
        xml_out.elem_text_esc("dc:title", &book.metadata.title)?;
    }
    if !book.metadata.description.is_empty() {
        xml_out.elem_text_esc("dc:description", &book.metadata.description)?;
    }
    if !book.metadata.subject.is_empty() {
        xml_out.elem_text_esc("dc:subject", &book.metadata.subject)?;
    }
    if !book.metadata.keyword.is_empty() {
        xml_out.elem_text_esc("meta:keyword", &book.metadata.keyword)?;
    }
    if !book.metadata.initial_creator.is_empty() {
        xml_out.elem_text_esc("meta:initial-creator", &book.metadata.initial_creator)?;
    }
    if !book.metadata.creator.is_empty() {
        xml_out.elem_text_esc("meta:creator", &book.metadata.creator)?;
    }
    if !book.metadata.printed_by.is_empty() {
        xml_out.elem_text_esc("meta:printed-by", &book.metadata.printed_by)?;
    }
    if let Some(v) = book.metadata.creation_date {
        xml_out.elem_text("meta:creation-date", &v.format(DATETIME_FORMAT))?;
    }
    if let Some(v) = book.metadata.date {
        xml_out.elem_text("dc:date", &v.format(DATETIME_FORMAT))?;
    }
    if let Some(v) = book.metadata.print_date {
        xml_out.elem_text("meta:print-date", &v.format(DATETIME_FORMAT))?;
    }
    if !book.metadata.language.is_empty() {
        xml_out.elem_text_esc("dc:language", &book.metadata.language)?;
    }
    if book.metadata.editing_cycles > 0 {
        xml_out.elem_text("meta:editing-cycles", &book.metadata.editing_cycles)?;
    }
    if book.metadata.editing_duration.num_seconds() > 0 {
        xml_out.elem_text(
            "meta:editing-duration",
            &format_duration2(book.metadata.editing_duration),
        )?;
    }

    if !book.metadata.template.is_empty() {
        xml_out.empty("meta:template")?;
        if let Some(v) = book.metadata.template.date {
            xml_out.attr("meta:date", &v.format(DATETIME_FORMAT))?;
        }
        if let Some(v) = book.metadata.template.actuate {
            xml_out.attr("xlink:actuate", &v)?;
        }
        if let Some(v) = &book.metadata.template.href {
            xml_out.attr_esc("xlink:href", v)?;
        }
        if let Some(v) = &book.metadata.template.title {
            xml_out.attr_esc("xlink:title", v)?;
        }
        if let Some(v) = book.metadata.template.link_type {
            xml_out.attr("xlink:type", &v)?;
        }
    }

    if !book.metadata.auto_reload.is_empty() {
        xml_out.empty("meta:auto_reload")?;
        if let Some(v) = book.metadata.auto_reload.delay {
            xml_out.attr("meta:delay", &format_duration2(v))?;
        }
        if let Some(v) = book.metadata.auto_reload.actuate {
            xml_out.attr("xlink:actuate", &v)?;
        }
        if let Some(v) = &book.metadata.auto_reload.href {
            xml_out.attr_esc("xlink:href", v)?;
        }
        if let Some(v) = &book.metadata.auto_reload.show {
            xml_out.attr("xlink:show", v)?;
        }
        if let Some(v) = book.metadata.auto_reload.link_type {
            xml_out.attr("xlink:type", &v)?;
        }
    }

    if !book.metadata.hyperlink_behaviour.is_empty() {
        xml_out.empty("meta:hyperlink-behaviour")?;
        if let Some(v) = &book.metadata.hyperlink_behaviour.target_frame_name {
            xml_out.attr_esc("office:target-frame-name", v)?;
        }
        if let Some(v) = &book.metadata.hyperlink_behaviour.show {
            xml_out.attr("xlink:show", v)?;
        }
    }

    xml_out.empty("meta:document-statistic")?;
    xml_out.attr(
        "meta:table-count",
        &book.metadata.document_statistics.table_count,
    )?;
    xml_out.attr(
        "meta:cell-count",
        &book.metadata.document_statistics.cell_count,
    )?;
    xml_out.attr(
        "meta:object-count",
        &book.metadata.document_statistics.object_count,
    )?;
    xml_out.attr(
        "meta:ole-object-count",
        &book.metadata.document_statistics.ole_object_count,
    )?;

    for userdef in &book.metadata.user_defined {
        xml_out.elem("meta:user-defined")?;
        xml_out.attr("meta:name", &userdef.name)?;
        if !matches!(userdef.value, MetaValue::String(_)) {
            xml_out.attr_str(
                "meta:value-type",
                match &userdef.value {
                    MetaValue::Boolean(_) => "boolean",
                    MetaValue::Datetime(_) => "date",
                    MetaValue::Float(_) => "float",
                    MetaValue::TimeDuration(_) => "time",
                    MetaValue::String(_) => unreachable!(),
                },
            )?;
        }
        match &userdef.value {
            MetaValue::Boolean(v) => xml_out.text_str(if *v { "true" } else { "false" })?,
            MetaValue::Datetime(v) => xml_out.text(&v.format(DATETIME_FORMAT))?,
            MetaValue::Float(v) => xml_out.text(&v)?,
            MetaValue::TimeDuration(v) => xml_out.text(&format_duration2(*v))?,
            MetaValue::String(v) => xml_out.text_esc(v)?,
        }
        xml_out.end_elem("meta:user-defined")?;
    }

    xml_out.end_elem("office:meta")?;
    Ok(())
}

fn write_ods_settings(book: &mut WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    let xmlns = book
        .xmlns
        .entry("settings.xml".into())
        .or_insert_with(NamespaceMap::new);

    xmlns.insert_str(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    );
    xmlns.insert_str("xmlns:ooo", "http://openoffice.org/2004/office");
    xmlns.insert_str(
        "xmlns:config",
        "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    );

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-settings")?;
    write_xmlns(xmlns, xml_out)?;
    xml_out.attr_esc("office:version", book.version())?;

    write_office_settings(book, xml_out)?;

    xml_out.end_elem("office:document-settings")?;

    xml_out.close()?;

    Ok(())
}

fn write_office_settings(book: &WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    xml_out.elem("office:settings")?;

    for (name, item) in book.config.iter() {
        match item {
            ConfigItem::Value(_) => {
                panic!("office-settings must not contain config-item");
            }
            ConfigItem::Set(_) => write_config_item_set(name, item, xml_out)?,
            ConfigItem::Vec(_) => {
                panic!("office-settings must not contain config-item-map-index")
            }
            ConfigItem::Map(_) => {
                panic!("office-settings must not contain config-item-map-named")
            }
            ConfigItem::Entry(_) => {
                panic!("office-settings must not contain config-item-map-entry")
            }
        }
    }

    xml_out.end_elem("office:settings")?;
    Ok(())
}

fn write_config_item_set(
    name: &str,
    set: &ConfigItem,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-set")?;
    xml_out.attr_esc("config:name", name)?;

    for (name, item) in set.iter() {
        match item {
            ConfigItem::Value(value) => write_config_item(name, value, xml_out)?,
            ConfigItem::Set(_) => write_config_item_set(name, item, xml_out)?,
            ConfigItem::Vec(_) => write_config_item_map_indexed(name, item, xml_out)?,
            ConfigItem::Map(_) => write_config_item_map_named(name, item, xml_out)?,
            ConfigItem::Entry(_) => {
                panic!("config-item-set must not contain config-item-map-entry")
            }
        }
    }

    xml_out.end_elem("config:config-item-set")?;

    Ok(())
}

fn write_config_item_map_indexed(
    name: &str,
    vec: &ConfigItem,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-map-indexed")?;
    xml_out.attr_esc("config:name", name)?;

    let mut index = 0;
    loop {
        let index_str = index.to_string();
        if let Some(item) = vec.get(&index_str) {
            match item {
                ConfigItem::Value(value) => write_config_item(name, value, xml_out)?,
                ConfigItem::Set(_) => {
                    panic!("config-item-map-index must not contain config-item-set")
                }
                ConfigItem::Vec(_) => {
                    panic!("config-item-map-index must not contain config-item-map-index")
                }
                ConfigItem::Map(_) => {
                    panic!("config-item-map-index must not contain config-item-map-named")
                }
                ConfigItem::Entry(_) => write_config_item_map_entry(None, item, xml_out)?,
            }
        } else {
            break;
        }

        index += 1;
    }

    xml_out.end_elem("config:config-item-map-indexed")?;

    Ok(())
}

fn write_config_item_map_named(
    name: &str,
    map: &ConfigItem,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-map-named")?;
    xml_out.attr_esc("config:name", name)?;

    for (name, item) in map.iter() {
        match item {
            ConfigItem::Value(value) => write_config_item(name, value, xml_out)?,
            ConfigItem::Set(_) => {
                panic!("config-item-map-index must not contain config-item-set")
            }
            ConfigItem::Vec(_) => {
                panic!("config-item-map-index must not contain config-item-map-index")
            }
            ConfigItem::Map(_) => {
                panic!("config-item-map-index must not contain config-item-map-named")
            }
            ConfigItem::Entry(_) => write_config_item_map_entry(Some(name), item, xml_out)?,
        }
    }

    xml_out.end_elem("config:config-item-map-named")?;

    Ok(())
}

fn write_config_item_map_entry(
    name: Option<&String>,
    map_entry: &ConfigItem,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-map-entry")?;
    if let Some(name) = name {
        xml_out.attr_esc("config:name", name)?;
    }

    for (name, item) in map_entry.iter() {
        match item {
            ConfigItem::Value(value) => write_config_item(name, value, xml_out)?,
            ConfigItem::Set(_) => write_config_item_set(name, item, xml_out)?,
            ConfigItem::Vec(_) => write_config_item_map_indexed(name, item, xml_out)?,
            ConfigItem::Map(_) => write_config_item_map_named(name, item, xml_out)?,
            ConfigItem::Entry(_) => {
                panic!("config:config-item-map-entry must not contain config-item-map-entry")
            }
        }
    }

    xml_out.end_elem("config:config-item-map-entry")?;

    Ok(())
}

fn write_config_item(
    name: &str,
    value: &ConfigValue,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    let is_empty = match value {
        ConfigValue::Base64Binary(t) => t.is_empty(),
        ConfigValue::String(t) => t.is_empty(),
        _ => false,
    };

    if is_empty {
        xml_out.empty("config:config-item")?;
    } else {
        xml_out.elem("config:config-item")?;
    }

    xml_out.attr_esc("config:name", name)?;

    match value {
        ConfigValue::Base64Binary(v) => {
            xml_out.attr_str("config:type", "base64Binary")?;
            xml_out.text(v)?;
        }
        ConfigValue::Boolean(v) => {
            xml_out.attr_str("config:type", "boolean")?;
            xml_out.text(&v)?;
        }
        ConfigValue::DateTime(v) => {
            xml_out.attr_str("config:type", "datetime")?;
            xml_out.text(&v.format(DATETIME_FORMAT))?;
        }
        ConfigValue::Double(v) => {
            xml_out.attr_str("config:type", "double")?;
            xml_out.text(&v)?;
        }
        ConfigValue::Int(v) => {
            xml_out.attr_str("config:type", "int")?;
            xml_out.text(&v)?;
        }
        ConfigValue::Long(v) => {
            xml_out.attr_str("config:type", "long")?;
            xml_out.text(&v)?;
        }
        ConfigValue::Short(v) => {
            xml_out.attr_str("config:type", "short")?;
            xml_out.text(&v)?;
        }
        ConfigValue::String(v) => {
            xml_out.attr_str("config:type", "string")?;
            xml_out.text(v)?;
        }
    }

    if !is_empty {
        xml_out.end_elem("config:config-item")?;
    }

    Ok(())
}

fn write_ods_styles(book: &mut WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    let xmlns = book
        .xmlns
        .entry("styles.xml".into())
        .or_insert_with(NamespaceMap::new);

    xmlns.insert_str(
        "xmlns:meta",
        "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    );
    xmlns.insert_str(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    );
    xmlns.insert_str(
        "xmlns:fo",
        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    );
    xmlns.insert_str("xmlns:ooo", "http://openoffice.org/2004/office");
    xmlns.insert_str("xmlns:xlink", "http://www.w3.org/1999/xlink");
    xmlns.insert_str("xmlns:dc", "http://purl.org/dc/elements/1.1/");
    xmlns.insert_str(
        "xmlns:style",
        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    );
    xmlns.insert_str(
        "xmlns:text",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    );
    xmlns.insert_str(
        "xmlns:dr3d",
        "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    );
    xmlns.insert_str(
        "xmlns:svg",
        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    );
    xmlns.insert_str(
        "xmlns:chart",
        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    );
    xmlns.insert_str("xmlns:rpt", "http://openoffice.org/2005/report");
    xmlns.insert_str(
        "xmlns:table",
        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    );
    xmlns.insert_str(
        "xmlns:number",
        "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    );
    xmlns.insert_str("xmlns:ooow", "http://openoffice.org/2004/writer");
    xmlns.insert_str("xmlns:oooc", "http://openoffice.org/2004/calc");
    xmlns.insert_str("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");
    xmlns.insert_str("xmlns:tableooo", "http://openoffice.org/2009/table");
    xmlns.insert_str(
        "xmlns:calcext",
        "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    );
    xmlns.insert_str("xmlns:drawooo", "http://openoffice.org/2010/draw");
    xmlns.insert_str(
        "xmlns:draw",
        "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    );
    xmlns.insert_str(
        "xmlns:loext",
        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    );
    xmlns.insert_str(
        "xmlns:field",
        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
    );
    xmlns.insert_str("xmlns:math", "http://www.w3.org/1998/Math/MathML");
    xmlns.insert_str(
        "xmlns:form",
        "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    );
    xmlns.insert_str(
        "xmlns:script",
        "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    );
    xmlns.insert_str("xmlns:dom", "http://www.w3.org/2001/xml-events");
    xmlns.insert_str("xmlns:xhtml", "http://www.w3.org/1999/xhtml");
    xmlns.insert_str("xmlns:grddl", "http://www.w3.org/2003/g/data-view#");
    xmlns.insert_str("xmlns:css3t", "http://www.w3.org/TR/css3-text/");
    xmlns.insert_str(
        "xmlns:presentation",
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    );

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-styles")?;
    write_xmlns(xmlns, xml_out)?;
    xml_out.attr_esc("office:version", book.version())?;

    write_office_font_face_decls(book, StyleOrigin::Styles, xml_out)?;
    write_office_styles(book, StyleOrigin::Styles, xml_out)?;
    write_office_automatic_styles(book, StyleOrigin::Styles, xml_out)?;
    write_office_master_styles(book, xml_out)?;

    xml_out.end_elem("office:document-styles")?;

    xml_out.close()?;

    Ok(())
}

fn write_ods_content(book: &mut WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    let xmlns = book
        .xmlns
        .entry("content.xml".into())
        .or_insert_with(NamespaceMap::new);

    xmlns.insert_str(
        "xmlns:meta",
        "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    );
    xmlns.insert_str(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    );
    xmlns.insert_str(
        "xmlns:fo",
        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    );
    xmlns.insert_str("xmlns:ooo", "http://openoffice.org/2004/office");
    xmlns.insert_str("xmlns:xlink", "http://www.w3.org/1999/xlink");
    xmlns.insert_str("xmlns:dc", "http://purl.org/dc/elements/1.1/");
    xmlns.insert_str(
        "xmlns:style",
        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    );
    xmlns.insert_str(
        "xmlns:text",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    );
    xmlns.insert_str(
        "xmlns:draw",
        "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    );
    xmlns.insert_str(
        "xmlns:dr3d",
        "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    );
    xmlns.insert_str(
        "xmlns:svg",
        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    );
    xmlns.insert_str(
        "xmlns:chart",
        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    );
    xmlns.insert_str("xmlns:rpt", "http://openoffice.org/2005/report");
    xmlns.insert_str(
        "xmlns:table",
        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    );
    xmlns.insert_str(
        "xmlns:number",
        "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    );
    xmlns.insert_str("xmlns:ooow", "http://openoffice.org/2004/writer");
    xmlns.insert_str("xmlns:oooc", "http://openoffice.org/2004/calc");
    xmlns.insert_str("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");
    xmlns.insert_str("xmlns:tableooo", "http://openoffice.org/2009/table");
    xmlns.insert_str(
        "xmlns:calcext",
        "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    );
    xmlns.insert_str("xmlns:drawooo", "http://openoffice.org/2010/draw");
    xmlns.insert_str(
        "xmlns:loext",
        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    );
    xmlns.insert_str(
        "xmlns:field",
        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
    );
    xmlns.insert_str("xmlns:math", "http://www.w3.org/1998/Math/MathML");
    xmlns.insert_str(
        "xmlns:form",
        "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    );
    xmlns.insert_str(
        "xmlns:script",
        "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    );
    xmlns.insert_str("xmlns:dom", "http://www.w3.org/2001/xml-events");
    xmlns.insert_str("xmlns:xforms", "http://www.w3.org/2002/xforms");
    xmlns.insert_str("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
    xmlns.insert_str("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
    xmlns.insert_str(
        "xmlns:formx",
        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
    );
    xmlns.insert_str("xmlns:xhtml", "http://www.w3.org/1999/xhtml");
    xmlns.insert_str("xmlns:grddl", "http://www.w3.org/2003/g/data-view#");
    xmlns.insert_str("xmlns:css3t", "http://www.w3.org/TR/css3-text/");
    xmlns.insert_str(
        "xmlns:presentation",
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    );

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-content")?;
    write_xmlns(xmlns, xml_out)?;
    xml_out.attr_esc("office:version", book.version())?;

    write_office_scripts(book, xml_out)?;
    write_office_font_face_decls(book, StyleOrigin::Content, xml_out)?;
    write_office_automatic_styles(book, StyleOrigin::Content, xml_out)?;

    write_office_body(book, xml_out)?;

    xml_out.end_elem("office:document-content")?;

    xml_out.close()?;

    Ok(())
}

fn write_office_body(book: &WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    xml_out.elem("office:body")?;
    xml_out.elem("office:spreadsheet")?;

    // extra tags. pass through only
    for tag in &book.extra {
        if tag.name() == "table:calculation-settings"
            || tag.name() == "table:label-ranges"
            || tag.name() == "table:tracked-changes"
            || tag.name() == "text:alphabetical-index-auto-mark-file"
            || tag.name() == "text:dde-connection-decls"
            || tag.name() == "text:sequence-decls"
            || tag.name() == "text:user-field-decls"
            || tag.name() == "text:variable-decls"
        {
            write_xmltag(tag, xml_out)?;
        }
    }

    write_content_validations(book, xml_out)?;

    for sheet in &book.sheets {
        write_sheet(book, sheet, xml_out)?;
    }

    // extra tags. pass through only
    for tag in &book.extra {
        if tag.name() == "table:consolidation"
            || tag.name() == "table:data-pilot-tables"
            || tag.name() == "table:database-ranges"
            || tag.name() == "table:dde-links"
            || tag.name() == "table:named-expressions"
            || tag.name() == "calcext:conditional-formats"
        {
            write_xmltag(tag, xml_out)?;
        }
    }

    xml_out.end_elem("office:spreadsheet")?;
    xml_out.end_elem("office:body")?;
    Ok(())
}

fn write_office_scripts(book: &WorkBook, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    xml_out.elem("office:scripts")?;
    write_scripts(&book.scripts, xml_out)?;
    write_event_listeners(&book.event_listener, xml_out)?;
    xml_out.end_elem("office:scripts")?;
    Ok(())
}

fn write_content_validations(
    book: &WorkBook,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    if !book.validations.is_empty() {
        xml_out.elem("table:content-validations")?;

        for valid in book.validations.values() {
            xml_out.elem("table:content-validation")?;
            xml_out.attr_esc("table:name", valid.name())?;
            xml_out.attr_esc("table:condition", &format_validation_condition(valid))?;
            xml_out.attr_str(
                "table:allow-empty-cell",
                if valid.allow_empty() { "true" } else { "false" },
            )?;
            xml_out.attr_str(
                "table:display-list",
                match valid.display() {
                    ValidationDisplay::NoDisplay => "no",
                    ValidationDisplay::Unsorted => "unsorted",
                    ValidationDisplay::SortAscending => "sort-ascending",
                },
            )?;
            xml_out.attr_esc("table:base-cell-address", &valid.base_cell())?;

            if let Some(err) = valid.err() {
                if err.text().is_some() {
                    xml_out.elem("table:error-message")?;
                } else {
                    xml_out.empty("table:error-message")?;
                }
                xml_out.attr("table:display", &err.display())?;
                xml_out.attr("table:message-type", &err.msg_type())?;
                if let Some(title) = err.title() {
                    xml_out.attr_esc("table:title", title)?;
                }
                if let Some(text) = err.text() {
                    write_xmltag(text, xml_out)?;
                }
                if err.text().is_some() {
                    xml_out.end_elem("table:error-message")?;
                }
            }
            if let Some(err) = valid.help() {
                if err.text().is_some() {
                    xml_out.elem("table:help-message")?;
                } else {
                    xml_out.empty("table:help-message")?;
                }
                xml_out.attr("table:display", &err.display())?;
                if let Some(title) = err.title() {
                    xml_out.attr_esc("table:title", title)?;
                }
                if let Some(text) = err.text() {
                    write_xmltag(text, xml_out)?;
                }
                if err.text().is_some() {
                    xml_out.end_elem("table:help-message")?;
                }
            }

            xml_out.end_elem("table:content-validation")?;
        }
        xml_out.end_elem("table:content-validations")?;
    }

    Ok(())
}

/// Is the cell hidden, and if yes how many more columns are hit.
fn check_hidden(ranges: &[CellRange], row: u32, col: u32) -> (bool, u32) {
    if let Some(found) = ranges.iter().find(|s| s.contains(row, col)) {
        (true, found.to_col() - col)
    } else {
        (false, 0)
    }
}

/// Removes any outlived Ranges from the vector.
fn remove_outlooped(ranges: &mut Vec<CellRange>, row: u32, col: u32) {
    *ranges = ranges
        .drain(..)
        .filter(|s| !s.out_looped(row, col))
        .collect();
}

fn write_sheet(
    book: &WorkBook,
    sheet: &Sheet,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("table:table")?;
    xml_out.attr_esc("table:name", &sheet.name)?;
    if let Some(style) = &sheet.style {
        xml_out.attr_esc("table:style-name", style)?;
    }
    if let Some(print_ranges) = &sheet.print_ranges {
        xml_out.attr_esc("table:print-ranges", &format_cellranges(print_ranges))?;
    }
    if !sheet.print() {
        xml_out.attr_str("table:print", "false")?;
    }
    if !sheet.display() {
        xml_out.attr_str("table:display", "false")?;
    }

    for tag in &sheet.extra {
        if tag.name() == "table:title"
            || tag.name() == "table:desc"
            || tag.name() == "table:table-source"
            || tag.name() == "office:dde-source"
            || tag.name() == "table:scenario"
            || tag.name() == "office:forms"
            || tag.name() == "table:shapes"
        {
            write_xmltag(tag, xml_out)?;
        }
    }

    let max_cell = sheet.used_grid_size();

    write_table_columns(sheet, max_cell, xml_out)?;

    // list of current spans
    let mut spans = Vec::<CellRange>::new();

    // table-row + table-cell
    let mut first_cell = true;
    let mut prev_row: u32 = 0;
    let mut prev_row_repeat: u32 = 1;
    let mut prev_col: u32 = 0;
    let mut row_group_count = 0;

    let mut it = sheet.into_iter();
    while let Some(((cur_row, cur_col), cell)) = it.next() {
        // Row repeat count.
        let cur_row_repeat = if let Some(row_header) = sheet.row_header.get(&cur_row) {
            row_header.repeat
        } else {
            1
        };

        // There may be a lot of gaps of any kind in our data.
        // In the XML format there is no cell identification, every gap
        // must be filled with empty rows/columns. For this we need some
        // calculations.

        // For the repeat-counter we need to look forward.
        let (next_row, next_col, is_last_cell) = //_
            if let Some((next_row, next_col)) = it.peek_cell()
        {
            (next_row, next_col, false)
        } else {
            (max_cell.0, max_cell.1, true)
        };

        // Looking forward row-wise.
        let forward_delta_row = next_row - cur_row;
        // Column deltas are only relevant in the same row, but we need to
        // fill up to max used columns.
        let forward_delta_col = if forward_delta_row >= 1 {
            max_cell.1 - cur_col
        } else {
            next_col - cur_col
        };

        // Looking backward row-wise.
        let backward_delta_row = cur_row - prev_row;
        // When a row changes our delta is from zero to cur_col.
        let backward_delta_col = if backward_delta_row >= 1 {
            cur_col
        } else {
            cur_col - prev_col
        };

        // After the first cell there is always an open row tag that
        // needs to be closed.
        if backward_delta_row > 0 && !first_cell {
            write_end_prev_row(
                sheet,
                prev_row,
                prev_row_repeat,
                &mut row_group_count,
                xml_out,
            )?;
        }

        // Any empty rows before this one?
        if backward_delta_row > 0 {
            // The repeat counter for the empty rows before is reduced by the last
            // real repeat.
            let mut synth_row_repeat = backward_delta_row as i32 - prev_row_repeat as i32;
            // At the very beginning last row is 0. But nothing has been written for it.
            // To account for this we add one repeat.
            if first_cell {
                synth_row_repeat += 1;
            }
            let synth_row = cur_row as i32 - synth_row_repeat;

            if synth_row_repeat > 0 {
                write_empty_rows_before(
                    sheet,
                    synth_row as u32,
                    synth_row_repeat as u32,
                    max_cell,
                    &mut row_group_count,
                    xml_out,
                )?;
            }
        }

        // Start a new row if there is a delta or we are at the start.
        // Fills in any blank cells before the current cell.
        if backward_delta_row > 0 || first_cell {
            write_start_current_row(
                sheet,
                cur_row,
                cur_row_repeat,
                backward_delta_col,
                &mut row_group_count,
                xml_out,
            )?;
        }

        // Remove no longer usefull cell-spans.
        remove_outlooped(&mut spans, cur_row, cur_col);

        // Current cell is hidden?
        let (is_hidden, hidden_cols) = check_hidden(&spans, cur_row, cur_col);

        // And now to something completely different ...
        write_cell(book, &cell, is_hidden, xml_out)?;

        // There may be some blank cells until the next one, but only one less the forward.
        if forward_delta_col > 1 {
            write_empty_cells(forward_delta_col, hidden_cols, xml_out)?;
        }

        // The last cell we will write? We can close the last row here,
        // where we have all the data.
        if is_last_cell {
            write_end_last_row(&mut row_group_count, xml_out)?;
        }

        // maybe span. only if visible, that nicely eliminates all
        // double hides.
        if let Some(span) = cell.span {
            if !is_hidden && (span.row_span > 1 || span.col_span > 1) {
                spans.push(CellRange::origin_span(cur_row, cur_col, span.into()));
            }
        }

        first_cell = false;
        prev_row = cur_row;
        prev_row_repeat = cur_row_repeat;
        prev_col = cur_col;
    }

    xml_out.end_elem("table:table")?;

    for tag in &sheet.extra {
        if tag.name() == "table:named-expressions" || tag.name() == "calcext:conditional-formats" {
            write_xmltag(tag, xml_out)?;
        }
    }

    Ok(())
}

fn write_empty_cells(
    mut forward_delta_col: u32,
    hidden_cols: u32,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    // split between hidden and regular cells.
    if hidden_cols >= forward_delta_col {
        xml_out.empty("covered-table-cell")?;
        xml_out.attr("table:number-columns-repeated", &(forward_delta_col - 1))?;

        forward_delta_col = 0;
    } else if hidden_cols > 0 {
        xml_out.empty("covered-table-cell")?;
        xml_out.attr("table:number-columns-repeated", &hidden_cols)?;

        forward_delta_col -= hidden_cols;
    }

    if forward_delta_col > 0 {
        xml_out.empty("table:table-cell")?;
        xml_out.attr("table:number-columns-repeated", &(forward_delta_col - 1))?;
    }

    Ok(())
}

fn write_start_current_row(
    sheet: &Sheet,
    cur_row: u32,
    cur_row_repeat: u32,
    backward_delta_col: u32,
    row_group_count: &mut u32,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    // groups
    for row_group in &sheet.group_rows {
        if row_group.from() == cur_row {
            *row_group_count += 1;
            xml_out.elem("table:table-row-group")?;
            if !row_group.display() {
                xml_out.attr_str("table:display", "false")?;
            }
        }
        // extra: if the group would start within our repeat range, it's started
        //        at the current row instead.
        if row_group.from() > cur_row && row_group.from() < cur_row + cur_row_repeat {
            *row_group_count += 1;
            xml_out.elem("table:table-row-group")?;
            if !row_group.display() {
                xml_out.attr_str("table:display", "false")?;
            }
        }
    }

    // row
    xml_out.elem("table:table-row")?;
    if let Some(row_header) = sheet.row_header.get(&cur_row) {
        if row_header.repeat > 1 {
            xml_out.attr_esc("table:number-rows-repeated", &row_header.repeat)?;
        }
        if let Some(rowstyle) = row_header.style() {
            xml_out.attr_esc("table:style-name", rowstyle)?;
        }
        if let Some(cellstyle) = row_header.cellstyle() {
            xml_out.attr_esc("table:default-cell-style-name", cellstyle)?;
        }
        if row_header.visible() != Visibility::Visible {
            xml_out.attr_esc("table:visibility", &row_header.visible())?;
        }
    }

    // Might not be the first column in this row.
    if backward_delta_col > 0 {
        xml_out.empty("table:table-cell")?;
        xml_out.attr_esc("table:number-columns-repeated", &backward_delta_col)?;
    }

    Ok(())
}

fn write_end_prev_row(
    sheet: &Sheet,
    last_r: u32,
    last_r_repeat: u32,
    row_group_count: &mut u32,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    // row
    xml_out.end_elem("table:table-row")?;

    // groups
    for row_group in &sheet.group_rows {
        if row_group.to() == last_r {
            *row_group_count -= 1;
            xml_out.end_elem("table:table-row-group")?;
        }
        // the group end is somewhere inside the repeated range.
        if row_group.to() > last_r && row_group.to() < last_r + last_r_repeat {
            *row_group_count -= 1;
            xml_out.end_elem("table:table-row-group")?;
        }
    }

    Ok(())
}

fn write_end_last_row(
    row_group_count: &mut u32,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    // row
    xml_out.end_elem("table:table-row")?;

    // close all groups
    while *row_group_count > 0 {
        *row_group_count -= 1;
        xml_out.end_elem("table:table-row-group")?;
    }

    Ok(())
}

fn write_empty_rows_before(
    sheet: &Sheet,
    last_row: u32,
    last_row_repeat: u32,
    max_cell: (u32, u32),
    row_group_count: &mut u32,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    // Are there any row groups? Then we don't use repeat but write everything out.
    if !sheet.group_rows.is_empty() {
        for r in last_row..last_row + last_row_repeat {
            // groups
            for row_group in &sheet.group_rows {
                if row_group.to() == r {
                    *row_group_count -= 1;
                    xml_out.end_elem("table:table-row-group")?;
                }
            }
            for row_group in &sheet.group_rows {
                if row_group.from() == r {
                    *row_group_count += 1;
                    xml_out.elem("table:table-row-group")?;
                    if !row_group.display() {
                        xml_out.attr_str("table:display", "false")?;
                    }
                }
            }
            // row
            write_empty_row(sheet, r, 1, max_cell, xml_out)?;
        }
    } else {
        write_empty_row(sheet, last_row, last_row_repeat, max_cell, xml_out)?;
    }

    Ok(())
}

fn write_empty_row(
    sheet: &Sheet,
    cur_row: u32,
    row_repeat: u32,
    max_cell: (u32, u32),
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("table:table-row")?;
    xml_out.attr("table:number-rows-repeated", &row_repeat)?;
    if let Some(row_header) = sheet.row_header.get(&cur_row) {
        if let Some(rowstyle) = row_header.style() {
            xml_out.attr_esc("table:style-name", rowstyle)?;
        }
        if let Some(cellstyle) = row_header.cellstyle() {
            xml_out.attr_esc("table:default-cell-style-name", cellstyle)?;
        }
        if row_header.visible() != Visibility::Visible {
            xml_out.attr_esc("table:visibility", &row_header.visible())?;
        }
    }

    // We fill the empty spaces completely up to max columns.
    xml_out.empty("table:table-cell")?;
    xml_out.attr("table:number-columns-repeated", &max_cell.1)?;

    xml_out.end_elem("table:table-row")?;

    Ok(())
}

fn write_table_columns(
    sheet: &Sheet,
    max_cell: (u32, u32),
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    // table:table-column
    for c in 0..max_cell.1 {
        for grp in &sheet.group_cols {
            if c == grp.from() {
                xml_out.elem("table:table-column-group")?;
                if !grp.display() {
                    xml_out.attr_str("table:display", "false")?;
                }
            }
        }

        xml_out.empty("table:table-column")?;
        if let Some(col_header) = sheet.col_header.get(&c) {
            if let Some(style) = col_header.style() {
                xml_out.attr_esc("table:style-name", style)?;
            }
            if let Some(cellstyle) = col_header.cellstyle() {
                xml_out.attr_esc("table:default-cell-style-name", cellstyle)?;
            }
            if col_header.visible() != Visibility::Visible {
                xml_out.attr_esc("table:visibility", &col_header.visible())?;
            }
        }

        for col_group in &sheet.group_cols {
            if c == col_group.to() {
                xml_out.end_elem("table:table-column-group")?;
            }
        }
    }

    Ok(())
}

#[allow(clippy::single_char_add_str)]
fn write_cell(
    book: &WorkBook,
    cell: &CellContentRef<'_>,
    is_hidden: bool,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    let tag = if is_hidden {
        "table:covered-table-cell"
    } else {
        "table:table-cell"
    };

    match cell.value {
        Value::Empty => xml_out.empty(tag)?,
        _ => xml_out.elem(tag)?,
    }

    if let Some(formula) = cell.formula {
        xml_out.attr_esc("table:formula", formula)?;
    }

    // Direct style oder value based default style.
    if let Some(style) = cell.style {
        xml_out.attr_esc("table:style-name", style)?;
    } else if let Some(style) = book.def_style(cell.value.value_type()) {
        xml_out.attr_esc("table:style-name", style)?;
    }

    // Content validation
    if let Some(validation_name) = cell.validation_name {
        xml_out.attr_esc("table:content-validation-name", validation_name)?;
    }

    // Spans
    if let Some(span) = cell.span {
        if span.row_span > 1 {
            xml_out.attr_esc("table:number-rows-spanned", &span.row_span)?;
        }
        if span.col_span > 1 {
            xml_out.attr_esc("table:number-columns-spanned", &span.col_span)?;
        }
    }

    // This finds the correct ValueFormat, but there is no way to use it.
    // Falls back to: Output the same string as needed for the value-attribute
    // and hope for the best. Seems to work well enough.
    //
    // let valuestyle = if let Some(style_name) = cell.style {
    //     book.find_value_format(style_name)
    // } else {
    //     None
    // };

    match cell.value {
        Value::Empty => {}
        Value::Text(s) => {
            xml_out.attr_str("office:value-type", "string")?;
            for l in s.split('\n') {
                xml_out.elem_text_esc("text:p", l)?;
            }
        }
        Value::TextXml(t) => {
            xml_out.attr_str("office:value-type", "string")?;
            for tt in t.iter() {
                write_xmltag(tt, xml_out)?;
            }
        }
        Value::DateTime(d) => {
            xml_out.attr_str("office:value-type", "date")?;
            let value = d.format(DATETIME_FORMAT);
            xml_out.attr("office:date-value", &value)?;
            xml_out.elem("text:p")?;
            xml_out.text(&value)?;
            xml_out.end_elem("text:p")?;
        }
        Value::TimeDuration(d) => {
            xml_out.attr_str("office:value-type", "time")?;
            let value = format_duration2(*d);
            xml_out.attr("office:time-value", &value)?;
            xml_out.elem("text:p")?;
            xml_out.text(&value)?;
            xml_out.end_elem("text:p")?;
        }
        Value::Boolean(b) => {
            xml_out.attr_str("office:value-type", "boolean")?;
            xml_out.attr_str("office:boolean-value", if *b { "true" } else { "false" })?;
            xml_out.elem("text:p")?;
            xml_out.text_str(if *b { "true" } else { "false" })?;
            xml_out.end_elem("text:p")?;
        }
        Value::Currency(v, c) => {
            xml_out.attr_str("office:value-type", "currency")?;
            xml_out.attr_esc("office:currency", c)?;
            xml_out.attr("office:value", v)?;
            xml_out.elem("text:p")?;
            xml_out.text_esc(c)?;
            xml_out.text_str(" ")?;
            xml_out.text(v)?;
            xml_out.end_elem("text:p")?;
        }
        Value::Number(v) => {
            xml_out.attr_str("office:value-type", "float")?;
            xml_out.attr("office:value", v)?;
            xml_out.elem("text:p")?;
            xml_out.text(v)?;
            xml_out.end_elem("text:p")?;
        }
        Value::Percentage(v) => {
            xml_out.attr_str("office:value-type", "percentage")?;
            xml_out.attr("office:value", v)?;
            xml_out.elem("text:p")?;
            xml_out.text(v)?;
            xml_out.end_elem("text:p")?;
        }
    }

    match cell.value {
        Value::Empty => {}
        _ => xml_out.end_elem(tag)?,
    }

    Ok(())
}

fn write_scripts(scripts: &Vec<Script>, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    for script in scripts {
        xml_out.elem("office:script")?;
        xml_out.attr_esc("script:language", &script.script_lang)?;

        for content in &script.script {
            write_xmlcontent(content, xml_out)?;
        }

        xml_out.end_elem("/office:script")?;
    }

    Ok(())
}

fn write_event_listeners(
    events: &HashMap<String, EventListener>,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    for event in events.values() {
        xml_out.empty("script:event-listener")?;
        xml_out.attr_esc("script:event-name", &event.event_name)?;
        xml_out.attr_esc("script:language", &event.script_lang)?;
        xml_out.attr_esc("script:macro-name", &event.macro_name)?;
        xml_out.attr_esc("xlink:actuate", &event.actuate)?;
        xml_out.attr_esc("xlink:href", &event.href)?;
        xml_out.attr_esc("xlink:type", &event.link_type)?;
    }

    Ok(())
}

fn write_office_font_face_decls(
    book: &WorkBook,
    origin: StyleOrigin,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("office:font-face-decls")?;
    write_style_font_face(&book.fonts, origin, xml_out)?;
    xml_out.end_elem("office:font-face-decls")?;
    Ok(())
}

fn write_style_font_face(
    fonts: &HashMap<String, FontFaceDecl>,
    origin: StyleOrigin,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    for font in fonts.values().filter(|s| s.origin() == origin) {
        xml_out.empty("style:font-face")?;
        xml_out.attr_esc("style:name", font.name())?;
        for (a, v) in font.attrmap().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    Ok(())
}

fn write_office_styles(
    book: &WorkBook,
    origin: StyleOrigin,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("office:styles")?;
    write_styles(book, origin, StyleUse::Default, xml_out)?;
    write_styles(book, origin, StyleUse::Named, xml_out)?;
    write_valuestyles(book, origin, StyleUse::Named, xml_out)?;
    write_valuestyles(book, origin, StyleUse::Default, xml_out)?;
    xml_out.end_elem("office:styles")?;
    Ok(())
}

fn write_office_automatic_styles(
    book: &WorkBook,
    origin: StyleOrigin,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("office:automatic-styles")?;
    write_pagestyles(&book.pagestyles, xml_out)?;
    write_styles(book, origin, StyleUse::Automatic, xml_out)?;
    write_valuestyles(book, origin, StyleUse::Automatic, xml_out)?;
    xml_out.end_elem("office:automatic-styles")?;
    Ok(())
}

fn write_office_master_styles(
    book: &WorkBook,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    xml_out.elem("office:master-styles")?;
    write_masterpage(&book.masterpages, xml_out)?;
    xml_out.end_elem("office:master-styles")?;
    Ok(())
}

fn write_styles(
    book: &WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    for style in book.colstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_colstyle(style, xml_out)?;
        }
    }
    for style in book.rowstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_rowstyle(style, xml_out)?;
        }
    }
    for style in book.tablestyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_tablestyle(style, xml_out)?;
        }
    }
    for style in book.cellstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_cellstyle(style, xml_out)?;
        }
    }
    for style in book.paragraphstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_paragraphstyle(style, xml_out)?;
        }
    }
    for style in book.textstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_textstyle(style, xml_out)?;
        }
    }
    for style in book.rubystyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_rubystyle(style, xml_out)?;
        }
    }
    for style in book.graphicstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_graphicstyle(style, xml_out)?;
        }
    }

    // if let Some(stylemaps) = style.stylemaps() {
    //     for sm in stylemaps {
    //         xml_out.empty("style:map")?;
    //         xml_out.attr_esc("style:condition", sm.condition())?;
    //         xml_out.attr_esc("style:apply-style-name", sm.applied_style())?;
    //         xml_out.attr_esc("style:base-cell-address", &sm.base_cell().to_string())?;
    //     }
    // }

    Ok(())
}

fn write_tablestyle(style: &TableStyle, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "table")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.tablestyle().is_empty() {
        xml_out.empty("style:table-properties")?;
        for (a, v) in style.tablestyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_rowstyle(style: &RowStyle, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "table-row")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.rowstyle().is_empty() {
        xml_out.empty("style:table-row-properties")?;
        for (a, v) in style.rowstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_colstyle(style: &ColStyle, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "table-column")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.colstyle().is_empty() {
        xml_out.empty("style:table-column-properties")?;
        for (a, v) in style.colstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_cellstyle(style: &CellStyle, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "table-cell")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.cellstyle().is_empty() {
        xml_out.empty("style:table-cell-properties")?;
        for (a, v) in style.cellstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if !&style.paragraphstyle().is_empty() {
        xml_out.empty("style:paragraph-properties")?;
        for (a, v) in style.paragraphstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if !style.textstyle().is_empty() {
        xml_out.empty("style:text-properties")?;
        for (a, v) in style.textstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if let Some(stylemaps) = style.stylemaps() {
        for sm in stylemaps {
            xml_out.empty("style:map")?;
            xml_out.attr_esc("style:condition", sm.condition())?;
            xml_out.attr_esc("style:apply-style-name", sm.applied_style())?;
            if let Some(r) = sm.base_cell() {
                xml_out.attr_esc("style:base-cell-address", r)?;
            }
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_paragraphstyle(
    style: &ParagraphStyle,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "paragraph")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.paragraphstyle().is_empty() {
        if style.tabstops().is_none() {
            xml_out.empty("style:paragraph-properties")?;
            for (a, v) in style.paragraphstyle().iter() {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        } else {
            xml_out.elem("style:paragraph-properties")?;
            for (a, v) in style.paragraphstyle().iter() {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
            xml_out.elem("style:tab-stops")?;
            if let Some(tabstops) = style.tabstops() {
                for ts in tabstops {
                    xml_out.empty("style:tab-stop")?;
                    for (a, v) in ts.attrmap().iter() {
                        xml_out.attr_esc(a.as_ref(), v)?;
                    }
                }
            }
            xml_out.end_elem("style:tab-stops")?;
            xml_out.end_elem("style:paragraph-properties")?;
        }
    }
    if !style.textstyle().is_empty() {
        xml_out.empty("style:text-properties")?;
        for (a, v) in style.textstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_textstyle(style: &TextStyle, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "text")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.textstyle().is_empty() {
        xml_out.empty("style:text-properties")?;
        for (a, v) in style.textstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_rubystyle(style: &RubyStyle, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "ruby")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.rubystyle().is_empty() {
        xml_out.empty("style:ruby-properties")?;
        for (a, v) in style.rubystyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_graphicstyle(
    style: &GraphicStyle,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr_str("style:family", "graphic")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }
    }

    if !style.graphicstyle().is_empty() {
        xml_out.empty("style:graphic-properties")?;
        for (a, v) in style.graphicstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }
    }

    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_valuestyles(
    book: &WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    write_valuestyle(&book.formats_boolean, origin, style_use, xml_out)?;
    write_valuestyle(&book.formats_currency, origin, style_use, xml_out)?;
    write_valuestyle(&book.formats_datetime, origin, style_use, xml_out)?;
    write_valuestyle(&book.formats_number, origin, style_use, xml_out)?;
    write_valuestyle(&book.formats_percentage, origin, style_use, xml_out)?;
    write_valuestyle(&book.formats_text, origin, style_use, xml_out)?;
    write_valuestyle(&book.formats_timeduration, origin, style_use, xml_out)?;
    Ok(())
}

fn write_valuestyle<T: ValueFormatTrait>(
    value_formats: &HashMap<String, T>,
    origin: StyleOrigin,
    styleuse: StyleUse,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    for value_format in value_formats
        .values()
        .filter(|s| s.origin() == origin && s.styleuse() == styleuse)
    {
        let tag = match value_format.value_type() {
            ValueType::Empty => unreachable!(),
            ValueType::Boolean => "number:boolean-style",
            ValueType::Number => "number:number-style",
            ValueType::Text => "number:text-style",
            ValueType::TextXml => "number:text-style",
            ValueType::TimeDuration => "number:time-style",
            ValueType::Percentage => "number:percentage-style",
            ValueType::Currency => "number:currency-style",
            ValueType::DateTime => "number:date-style",
        };

        xml_out.elem(tag)?;
        xml_out.attr_esc("style:name", value_format.name())?;
        for (a, v) in value_format.attrmap().iter() {
            xml_out.attr_esc(a.as_ref(), v)?;
        }

        if !value_format.textstyle().is_empty() {
            xml_out.empty("style:text-properties")?;
            for (a, v) in value_format.textstyle().iter() {
                xml_out.attr_esc(a.as_ref(), v)?;
            }
        }

        for part in value_format.parts() {
            let part_tag = match part.part_type() {
                FormatPartType::Boolean => "number:boolean",
                FormatPartType::Number => "number:number",
                FormatPartType::ScientificNumber => "number:scientific-number",
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
                FormatPartType::Text => "number:text",
                FormatPartType::TextContent => "number:text-content",
                FormatPartType::FillCharacter => "number:fill-character",
            };

            if part.part_type() == FormatPartType::Text
                || part.part_type() == FormatPartType::CurrencySymbol
                || part.part_type() == FormatPartType::FillCharacter
            {
                xml_out.elem(part_tag)?;
                for (a, v) in part.attrmap().iter() {
                    xml_out.attr_esc(a.as_ref(), v)?;
                }
                if let Some(content) = part.content() {
                    xml_out.text_esc(content)?;
                }
                xml_out.end_elem(part_tag)?;
            } else if part.part_type() == FormatPartType::Number {
                if let Some(embedded_text) = part.content() {
                    xml_out.elem(part_tag)?;
                    for (a, v) in part.attrmap().iter() {
                        xml_out.attr_esc(a.as_ref(), v)?;
                    }

                    // embedded text
                    xml_out.elem("number:embedded-text")?;
                    xml_out.attr_esc("number:position", &part.position())?;
                    xml_out.text_esc(embedded_text)?;
                    xml_out.end_elem("number:embedded-text")?;

                    xml_out.end_elem(part_tag)?;
                } else {
                    xml_out.empty(part_tag)?;
                    for (a, v) in part.attrmap().iter() {
                        xml_out.attr_esc(a.as_ref(), v)?;
                    }
                }
            } else {
                xml_out.empty(part_tag)?;
                for (a, v) in part.attrmap().iter() {
                    xml_out.attr_esc(a.as_ref(), v)?;
                }
            }
        }

        if let Some(stylemaps) = value_format.stylemaps() {
            for sm in stylemaps {
                xml_out.empty("style:map")?;
                xml_out.attr_esc("style:condition", sm.condition())?;
                xml_out.attr_esc("style:apply-style-name", sm.applied_style())?;
            }
        }

        xml_out.end_elem(tag)?;
    }

    Ok(())
}

fn write_pagestyles(
    styles: &HashMap<String, PageStyle>,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    for style in styles.values() {
        xml_out.elem("style:page-layout")?;
        xml_out.attr_esc("style:name", style.name())?;
        if let Some(master_page_usage) = &style.master_page_usage {
            xml_out.attr_esc("style:page-usage", master_page_usage)?;
        }

        if !style.style().is_empty() {
            xml_out.empty("style:page-layout-properties")?;
            for (k, v) in style.style().iter() {
                xml_out.attr_esc(k.as_ref(), v)?;
            }
        }

        xml_out.elem("style:header-style")?;
        xml_out.empty("style:header-footer-properties")?;
        if !style.headerstyle().style().is_empty() {
            for (k, v) in style.headerstyle().style().iter() {
                xml_out.attr_esc(k.as_ref(), v)?;
            }
        }
        xml_out.end_elem("style:header-style")?;

        xml_out.elem("style:footer-style")?;
        xml_out.empty("style:header-footer-properties")?;
        if !style.footerstyle().style().is_empty() {
            for (k, v) in style.footerstyle().style().iter() {
                xml_out.attr_esc(k.as_ref(), v)?;
            }
        }
        xml_out.end_elem("style:footer-style")?;

        xml_out.end_elem("style:page-layout")?;
    }

    Ok(())
}

fn write_masterpage(
    styles: &HashMap<String, MasterPage>,
    xml_out: &mut OdsXmlWriter<'_>,
) -> Result<(), OdsError> {
    for style in styles.values() {
        xml_out.elem("style:master-page")?;
        xml_out.attr_esc("style:name", style.name())?;
        if !style.display_name().is_empty() {
            xml_out.attr_esc("style:display-name", style.display_name())?;
        }
        xml_out.attr_esc("style:page-layout-name", style.pagestyle())?;

        xml_out.elem("style:header")?;
        if !style.header().display() {
            xml_out.attr_str("style:display", "false")?;
        }
        write_regions(style.header(), xml_out)?;
        xml_out.end_elem("style:header")?;

        if !style.header_first().is_empty() {
            xml_out.elem("style:header-first")?;
            if !style.header_first().display() {
                xml_out.attr_str("style:display", "false")?;
            }
            write_regions(style.header_first(), xml_out)?;
            xml_out.end_elem("style:header-first")?;
        }

        xml_out.elem("style:header-left")?;
        if !style.header_left().display() || style.header_left().is_empty() {
            xml_out.attr_str("style:display", "false")?;
        }
        write_regions(style.header_left(), xml_out)?;
        xml_out.end_elem("style:header-left")?;

        xml_out.elem("style:footer")?;
        if !style.footer().display() {
            xml_out.attr_str("style:display", "false")?;
        }
        write_regions(style.footer(), xml_out)?;
        xml_out.end_elem("style:footer")?;

        if !style.footer_first().is_empty() {
            xml_out.elem("style:footer-first")?;
            if !style.footer_first().display() {
                xml_out.attr_str("style:display", "false")?;
            }
            write_regions(style.footer_first(), xml_out)?;
            xml_out.end_elem("style:footer-first")?;
        }

        xml_out.elem("style:footer-left")?;
        if !style.footer_left().display() || style.footer_left().is_empty() {
            xml_out.attr_str("style:display", "false")?;
        }
        write_regions(style.footer_left(), xml_out)?;
        xml_out.end_elem("style:footer-left")?;

        xml_out.end_elem("style:master-page")?;
    }

    Ok(())
}

fn write_regions(hf: &HeaderFooter, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    for left in hf.left() {
        xml_out.elem("style:region-left")?;
        write_xmltag(left, xml_out)?;
        xml_out.end_elem("style:region-left")?;
    }
    for center in hf.center() {
        xml_out.elem("style:region-center")?;
        write_xmltag(center, xml_out)?;
        xml_out.end_elem("style:region-center")?;
    }
    for right in hf.right() {
        xml_out.elem("style:region-right")?;
        write_xmltag(right, xml_out)?;
        xml_out.end_elem("style:region-right")?;
    }
    for content in hf.content() {
        write_xmltag(content, xml_out)?;
    }

    Ok(())
}

fn write_xmlcontent(x: &XmlContent, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    match x {
        XmlContent::Text(t) => {
            xml_out.text_esc(t)?;
        }
        XmlContent::Tag(t) => {
            write_xmltag(t, xml_out)?;
        }
    }
    Ok(())
}

fn write_xmltag(x: &XmlTag, xml_out: &mut OdsXmlWriter<'_>) -> Result<(), OdsError> {
    if x.is_empty() {
        xml_out.empty(x.name())?;
    } else {
        xml_out.elem(x.name())?;
    }
    for (k, v) in x.attrmap().iter() {
        xml_out.attr_esc(k.as_ref(), v)?;
    }

    for c in x.content() {
        match c {
            XmlContent::Text(t) => {
                xml_out.text_esc(t)?;
            }
            XmlContent::Tag(t) => {
                write_xmltag(t, xml_out)?;
            }
        }
    }

    if !x.is_empty() {
        xml_out.end_elem(x.name())?;
    }

    Ok(())
}

// All extra entries from the manifest.
fn write_ods_extra<W: Write + Seek>(
    book: &WorkBook,
    zip_writer: &mut ZipOut<W>,
) -> Result<(), OdsError> {
    for manifest in book.manifest.values() {
        if !matches!(
            manifest.full_path.as_str(),
            "/" | "settings.xml" | "styles.xml" | "content.xml" | "meta.xml"
        ) {
            if manifest.is_dir() {
                zip_writer.add_directory(&manifest.full_path, FileOptions::default())?;
            } else {
                let mut wr = zip_writer.start_file(&manifest.full_path, FileOptions::default())?;
                if let Some(buf) = &manifest.buffer {
                    wr.write_all(buf.as_slice())?;
                }
            }
        }
    }

    Ok(())
}
