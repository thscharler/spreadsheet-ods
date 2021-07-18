use std::collections::HashMap;
use std::fs::File;
use std::io;
use std::io::{Cursor, Seek, Write};
use std::path::Path;

use chrono::NaiveDateTime;
use zip::write::FileOptions;

use crate::config::{ConfigItem, ConfigItemType, ConfigValue};
use crate::error::OdsError;
use crate::format::FormatPartType;
use crate::io::xmlwriter::XmlWriter;
use crate::io::zip_out::{ZipOut, ZipWrite};
use crate::io::FileBufEntry;
use crate::refs::{cellranges_string, CellRange};
use crate::style::{
    CellStyle, ColStyle, FontFaceDecl, GraphicStyle, HeaderFooter, MasterPage, PageStyle,
    ParagraphStyle, RowStyle, StyleOrigin, StyleUse, TableStyle, TextStyle,
};
use crate::validation::ValidationDisplay;
use crate::xmltree::{XmlContent, XmlTag};
use crate::{ucell, CellData, Length, Sheet, Value, ValueFormat, ValueType, Visibility, WorkBook};

type OdsWriter<W> = ZipOut<W>;
type XmlOdsWriter<'a, W> = XmlWriter<ZipWrite<'a, W>>;

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

/// Writes the ODS file.
///
/// All the parts are written to a temp directory and then zipped together.
///
pub fn write_ods<P: AsRef<Path>>(book: &mut WorkBook, ods_path: P) -> Result<(), OdsError> {
    let zip_writer = ZipOut::<File>::new_file(ods_path.as_ref())?;
    write_ods_impl(book, zip_writer)?;
    Ok(())
}

/// Writes the ODS file.
///
/// All the parts are written to a temp directory and then zipped together.
///
fn write_ods_impl<W: Write + Seek>(
    book: &mut WorkBook,
    mut zip_writer: OdsWriter<W>,
) -> Result<W, OdsError> {
    sanity_checks(book)?;

    store_derived(book)?;

    // copy all buffered data from the original.
    copy_workbook(&book, &mut zip_writer)?;
    // write the rest, if necessary.
    write_mimetype(&book, &mut zip_writer)?;
    write_manifest(&book, &mut zip_writer)?;
    write_manifest_rdf(&book, &mut zip_writer)?;
    write_meta(&book, &mut zip_writer)?;
    // not in use any more, just ignore
    // write_configurations(&mut zip_writer, &mut file_set)?;
    write_settings(&book, &mut zip_writer)?;
    write_ods_styles(&book, &mut zip_writer)?;
    write_ods_content(&book, &mut zip_writer)?;

    Ok(zip_writer.zip()?)
}

fn sanity_checks(book: &mut WorkBook) -> Result<(), OdsError> {
    if book.sheets.is_empty() {
        return Err(OdsError::Ods("Workbook contains no sheets.".to_string()));
    }
    Ok(())
}

#[allow(clippy::collapsible_else_if)]
#[allow(clippy::collapsible_if)]
fn store_derived(book: &mut WorkBook) -> Result<(), OdsError> {
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
                    let colstyle = book.add_colstyle(ColStyle::empty());
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
                    let rowstyle = book.add_rowstyle(RowStyle::empty());
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
        bc.insert("CodeName", sheet.name().as_str().to_string());

        book.attach_sheet(sheet);
    }

    book.config.attach(config);

    Ok(())
}

fn copy_workbook<W: Write + Seek>(
    book: &WorkBook,
    zip_writer: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    for filebuf in book.filebuf.iter() {
        match filebuf {
            FileBufEntry::Dir(name) => {
                zip_writer.add_directory(name, FileOptions::default())?;
            }
            FileBufEntry::File(name, buf) => {
                let mut wr = zip_writer.start_file(name, FileOptions::default())?;
                wr.write_all(buf.as_slice())?;
            }
        }
    }

    Ok(())
}

fn write_mimetype<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), io::Error> {
    if !book.filebuf.contains("mimetype") {
        let mut w = zip_out.start_file(
            "mimetype",
            FileOptions::default().compression_method(zip::CompressionMethod::Stored),
        )?;

        let mime = "application/vnd.oasis.opendocument.spreadsheet";
        w.write_all(mime.as_bytes())?;
    }

    Ok(())
}

fn write_manifest<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    if !book.filebuf.contains("META-INF/manifest.xml") {
        zip_out.add_directory("META-INF", FileOptions::default())?;
        let w = zip_out.start_file("META-INF/manifest.xml", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("manifest:manifest")?;
        xml_out.attr(
            "xmlns:manifest",
            "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
        )?;
        xml_out.attr("manifest:version", book.version())?;

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "/")?;
        xml_out.attr("manifest:version", book.version())?;
        xml_out.attr(
            "manifest:media-type",
            "application/vnd.oasis.opendocument.spreadsheet",
        )?;

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

        xml_out.empty("manifest:file-entry")?;
        xml_out.attr("manifest:full-path", "settings.xml")?;
        xml_out.attr("manifest:media-type", "text/xml")?;

        xml_out.end_elem("manifest:manifest")?;

        xml_out.close()?;
    }

    Ok(())
}

fn write_manifest_rdf<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    if !book.filebuf.contains("manifest.rdf") {
        let w = zip_out.start_file("manifest.rdf", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("rdf:RDF")?;
        xml_out.attr("xmlns:rdf", "http://www.w3.org/1999/02/22-rdf-syntax-ns#")?;

        xml_out.elem("rdf:Description")?;
        xml_out.attr("rdf:about", "content.xml")?;

        xml_out.empty("rdf:type")?;
        xml_out.attr(
            "rdf:resource",
            "http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile",
        )?;

        xml_out.end_elem("rdf:Description")?;

        xml_out.elem("rdf:Description")?;
        xml_out.attr("rdf:about", "")?;

        xml_out.empty("ns0:hasPart")?;
        xml_out.attr(
            "xmlns:ns0",
            "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#",
        )?;
        xml_out.attr("rdf:resource", "content.xml")?;

        xml_out.end_elem("rdf:Description")?;

        xml_out.elem("rdf:Description")?;
        xml_out.attr("rdf:about", "")?;

        xml_out.empty("rdf:type")?;
        xml_out.attr(
            "rdf:resource",
            "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document",
        )?;

        xml_out.end_elem("rdf:Description")?;

        xml_out.end_elem("rdf:RDF")?;

        xml_out.close()?;
    }

    Ok(())
}

fn write_meta<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    if !book.filebuf.contains("meta.xml") {
        let w = zip_out.start_file("meta.xml", FileOptions::default())?;

        let mut xml_out = XmlWriter::new(w);

        xml_out.dtd("UTF-8")?;

        xml_out.elem("office:document-meta")?;
        xml_out.attr(
            "xmlns:meta",
            "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
        )?;
        xml_out.attr(
            "xmlns:office",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        )?;
        xml_out.attr("office:version", book.version())?;

        xml_out.elem("office:meta")?;

        xml_out.elem_text("meta:generator", "spreadsheet-ods 0.8.1")?;
        let s = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?;
        let d = NaiveDateTime::from_timestamp(s.as_secs() as i64, 0);
        xml_out.elem_text(
            "meta:creation-date",
            &d.format("%Y-%m-%dT%H:%M:%S%.f").to_string(),
        )?;
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

fn write_settings<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    let w = zip_out.start_file("settings.xml", FileOptions::default())?;

    let mut xml_out = XmlWriter::new(w);

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-settings")?;
    xml_out.attr(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    )?;
    xml_out.attr("xmlns:ooo", "http://openoffice.org/2004/office")?;
    xml_out.attr(
        "xmlns:config",
        "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    )?;
    xml_out.attr("office:version", book.version())?;
    xml_out.elem("office:settings")?;

    for (name, item) in book.config.iter() {
        match item {
            ConfigItem::Value(_) => {
                panic!("office-settings must not contain config-item");
            }
            ConfigItem::Set(_) => write_config_item_set(name, item, &mut xml_out)?,
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
    xml_out.end_elem("office:document-settings")?;

    xml_out.close()?;

    Ok(())
}

fn write_config_item_set<W: Write + Seek>(
    name: &str,
    set: &ConfigItem,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-set")?;
    xml_out.attr("config:name", name)?;

    for (name, item) in set.iter() {
        match item {
            ConfigItem::Value(value) => write_config_item(name, &value, xml_out)?,
            ConfigItem::Set(_) => write_config_item_set(name, &item, xml_out)?,
            ConfigItem::Vec(_) => write_config_item_map_indexed(name, &item, xml_out)?,
            ConfigItem::Map(_) => write_config_item_map_named(name, &item, xml_out)?,
            ConfigItem::Entry(_) => {
                panic!("config-item-set must not contain config-item-map-entry")
            }
        }
    }

    xml_out.end_elem("config:config-item-set")?;

    Ok(())
}

fn write_config_item_map_indexed<W: Write + Seek>(
    name: &str,
    vec: &ConfigItem,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-map-indexed")?;
    xml_out.attr("config:name", name)?;

    let mut index = 0;
    loop {
        let index_str = index.to_string();
        if let Some(item) = vec.get(&index_str) {
            match item {
                ConfigItem::Value(value) => write_config_item(name, &value, xml_out)?,
                ConfigItem::Set(_) => {
                    panic!("config-item-map-index must not contain config-item-set")
                }
                ConfigItem::Vec(_) => {
                    panic!("config-item-map-index must not contain config-item-map-index")
                }
                ConfigItem::Map(_) => {
                    panic!("config-item-map-index must not contain config-item-map-named")
                }
                ConfigItem::Entry(_) => write_config_item_map_entry(None, &item, xml_out)?,
            }
        } else {
            break;
        }

        index += 1;
    }

    xml_out.end_elem("config:config-item-map-indexed")?;

    Ok(())
}

fn write_config_item_map_named<W: Write + Seek>(
    name: &str,
    map: &ConfigItem,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-map-named")?;
    xml_out.attr("config:name", name)?;

    for (name, item) in map.iter() {
        match item {
            ConfigItem::Value(value) => write_config_item(name, &value, xml_out)?,
            ConfigItem::Set(_) => {
                panic!("config-item-map-index must not contain config-item-set")
            }
            ConfigItem::Vec(_) => {
                panic!("config-item-map-index must not contain config-item-map-index")
            }
            ConfigItem::Map(_) => {
                panic!("config-item-map-index must not contain config-item-map-named")
            }
            ConfigItem::Entry(_) => write_config_item_map_entry(Some(name), &item, xml_out)?,
        }
    }

    xml_out.end_elem("config:config-item-map-named")?;

    Ok(())
}

fn write_config_item_map_entry<W: Write + Seek>(
    name: Option<&String>,
    map_entry: &ConfigItem,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.elem("config:config-item-map-entry")?;
    if let Some(name) = name {
        xml_out.attr("config:name", name)?;
    }

    for (name, item) in map_entry.iter() {
        match item {
            ConfigItem::Value(value) => write_config_item(name, &value, xml_out)?,
            ConfigItem::Set(_) => write_config_item_set(name, &item, xml_out)?,
            ConfigItem::Vec(_) => write_config_item_map_indexed(name, &item, xml_out)?,
            ConfigItem::Map(_) => write_config_item_map_named(name, &item, xml_out)?,
            ConfigItem::Entry(_) => {
                panic!("config:config-item-map-entry must not contain config-item-map-entry")
            }
        }
    }

    xml_out.end_elem("config:config-item-map-entry")?;

    Ok(())
}

fn write_config_item<W: Write + Seek>(
    name: &str,
    value: &ConfigValue,
    xml_out: &mut XmlOdsWriter<W>,
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

    xml_out.attr("config:name", name)?;

    match value {
        ConfigValue::Base64Binary(v) => {
            xml_out.attr("config:type", "base64Binary")?;
            xml_out.text(v)?;
        }
        ConfigValue::Boolean(v) => {
            xml_out.attr("config:type", "boolean")?;
            xml_out.text(&v.to_string())?;
        }
        ConfigValue::DateTime(v) => {
            xml_out.attr("config:type", "datetime")?;
            xml_out.text(&v.format("%Y-%m-%dT%H:%M:%S%.f").to_string())?;
        }
        ConfigValue::Double(v) => {
            xml_out.attr("config:type", "double")?;
            xml_out.text(&v.to_string())?;
        }
        ConfigValue::Int(v) => {
            xml_out.attr("config:type", "int")?;
            xml_out.text(&v.to_string())?;
        }
        ConfigValue::Long(v) => {
            xml_out.attr("config:type", "long")?;
            xml_out.text(&v.to_string())?;
        }
        ConfigValue::Short(v) => {
            xml_out.attr("config:type", "short")?;
            xml_out.text(&v.to_string())?;
        }
        ConfigValue::String(v) => {
            xml_out.attr("config:type", "string")?;
            xml_out.text(v)?;
        }
    }

    if !is_empty {
        xml_out.end_elem("config:config-item")?;
    }

    Ok(())
}

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

fn write_ods_styles<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    let w = zip_out.start_file("styles.xml", FileOptions::default())?;

    let mut xml_out = XmlWriter::new(w);

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-styles")?;
    xml_out.attr(
        "xmlns:meta",
        "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    )?;
    xml_out.attr(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    )?;
    xml_out.attr(
        "xmlns:fo",
        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    )?;
    xml_out.attr("xmlns:ooo", "http://openoffice.org/2004/office")?;
    xml_out.attr("xmlns:xlink", "http://www.w3.org/1999/xlink")?;
    xml_out.attr("xmlns:dc", "http://purl.org/dc/elements/1.1/")?;
    xml_out.attr(
        "xmlns:style",
        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    )?;
    xml_out.attr(
        "xmlns:text",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    )?;
    xml_out.attr(
        "xmlns:dr3d",
        "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    )?;
    xml_out.attr(
        "xmlns:svg",
        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    )?;
    xml_out.attr(
        "xmlns:chart",
        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    )?;
    xml_out.attr("xmlns:rpt", "http://openoffice.org/2005/report")?;
    xml_out.attr(
        "xmlns:table",
        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    )?;
    xml_out.attr(
        "xmlns:number",
        "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    )?;
    xml_out.attr("xmlns:ooow", "http://openoffice.org/2004/writer")?;
    xml_out.attr("xmlns:oooc", "http://openoffice.org/2004/calc")?;
    xml_out.attr("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")?;
    xml_out.attr("xmlns:tableooo", "http://openoffice.org/2009/table")?;
    xml_out.attr(
        "xmlns:calcext",
        "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    )?;
    xml_out.attr("xmlns:drawooo", "http://openoffice.org/2010/draw")?;
    xml_out.attr(
        "xmlns:draw",
        "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    )?;
    xml_out.attr(
        "xmlns:loext",
        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    )?;
    xml_out.attr(
        "xmlns:field",
        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
    )?;
    xml_out.attr("xmlns:math", "http://www.w3.org/1998/Math/MathML")?;
    xml_out.attr(
        "xmlns:form",
        "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    )?;
    xml_out.attr(
        "xmlns:script",
        "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    )?;
    xml_out.attr("xmlns:dom", "http://www.w3.org/2001/xml-events")?;
    xml_out.attr("xmlns:xhtml", "http://www.w3.org/1999/xhtml")?;
    xml_out.attr("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")?;
    xml_out.attr("xmlns:css3t", "http://www.w3.org/TR/css3-text/")?;
    xml_out.attr(
        "xmlns:presentation",
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    )?;
    xml_out.attr("office:version", book.version())?;

    xml_out.elem("office:font-face-decls")?;
    write_font_decl(&book.fonts, StyleOrigin::Styles, &mut xml_out)?;
    xml_out.end_elem("office:font-face-decls")?;

    xml_out.elem("office:styles")?;
    write_styles(book, StyleOrigin::Styles, StyleUse::Default, &mut xml_out)?;
    write_styles(book, StyleOrigin::Styles, StyleUse::Named, &mut xml_out)?;
    write_valuestyles(
        &book.formats,
        StyleOrigin::Styles,
        StyleUse::Named,
        &mut xml_out,
    )?;
    write_valuestyles(
        &book.formats,
        StyleOrigin::Styles,
        StyleUse::Default,
        &mut xml_out,
    )?;
    xml_out.end_elem("office:styles")?;

    xml_out.elem("office:automatic-styles")?;
    write_pagestyles(&book.pagestyles, &mut xml_out)?;
    write_styles(book, StyleOrigin::Styles, StyleUse::Automatic, &mut xml_out)?;
    write_valuestyles(
        &book.formats,
        StyleOrigin::Styles,
        StyleUse::Automatic,
        &mut xml_out,
    )?;
    xml_out.end_elem("office:automatic-styles")?;

    xml_out.elem("office:master-styles")?;
    write_masterpage(&book.masterpages, &mut xml_out)?;
    xml_out.end_elem("office:master-styles")?;

    xml_out.end_elem("office:document-styles")?;

    xml_out.close()?;

    Ok(())
}

fn write_ods_content<W: Write + Seek>(
    book: &WorkBook,
    zip_out: &mut OdsWriter<W>,
) -> Result<(), OdsError> {
    let w = zip_out.start_file("content.xml", FileOptions::default())?;
    let mut xml_out = XmlWriter::new(w);

    xml_out.dtd("UTF-8")?;

    xml_out.elem("office:document-content")?;
    xml_out.attr(
        "xmlns:meta",
        "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    )?;
    xml_out.attr(
        "xmlns:office",
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    )?;
    xml_out.attr(
        "xmlns:fo",
        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    )?;
    xml_out.attr("xmlns:ooo", "http://openoffice.org/2004/office")?;
    xml_out.attr("xmlns:xlink", "http://www.w3.org/1999/xlink")?;
    xml_out.attr("xmlns:dc", "http://purl.org/dc/elements/1.1/")?;
    xml_out.attr(
        "xmlns:style",
        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    )?;
    xml_out.attr(
        "xmlns:text",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    )?;
    xml_out.attr(
        "xmlns:draw",
        "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    )?;
    xml_out.attr(
        "xmlns:dr3d",
        "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    )?;
    xml_out.attr(
        "xmlns:svg",
        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    )?;
    xml_out.attr(
        "xmlns:chart",
        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    )?;
    xml_out.attr("xmlns:rpt", "http://openoffice.org/2005/report")?;
    xml_out.attr(
        "xmlns:table",
        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    )?;
    xml_out.attr(
        "xmlns:number",
        "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    )?;
    xml_out.attr("xmlns:ooow", "http://openoffice.org/2004/writer")?;
    xml_out.attr("xmlns:oooc", "http://openoffice.org/2004/calc")?;
    xml_out.attr("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")?;
    xml_out.attr("xmlns:tableooo", "http://openoffice.org/2009/table")?;
    xml_out.attr(
        "xmlns:calcext",
        "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    )?;
    xml_out.attr("xmlns:drawooo", "http://openoffice.org/2010/draw")?;
    xml_out.attr(
        "xmlns:loext",
        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    )?;
    xml_out.attr(
        "xmlns:field",
        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
    )?;
    xml_out.attr("xmlns:math", "http://www.w3.org/1998/Math/MathML")?;
    xml_out.attr(
        "xmlns:form",
        "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    )?;
    xml_out.attr(
        "xmlns:script",
        "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    )?;
    xml_out.attr("xmlns:dom", "http://www.w3.org/2001/xml-events")?;
    xml_out.attr("xmlns:xforms", "http://www.w3.org/2002/xforms")?;
    xml_out.attr("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")?;
    xml_out.attr("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")?;
    xml_out.attr(
        "xmlns:formx",
        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
    )?;
    xml_out.attr("xmlns:xhtml", "http://www.w3.org/1999/xhtml")?;
    xml_out.attr("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")?;
    xml_out.attr("xmlns:css3t", "http://www.w3.org/TR/css3-text/")?;
    xml_out.attr(
        "xmlns:presentation",
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    )?;

    xml_out.attr("office:version", book.version())?;

    xml_out.empty("office:scripts")?;

    xml_out.elem("office:font-face-decls")?;
    write_font_decl(&book.fonts, StyleOrigin::Content, &mut xml_out)?;
    xml_out.end_elem("office:font-face-decls")?;

    xml_out.elem("office:automatic-styles")?;
    write_styles(
        book,
        StyleOrigin::Content,
        StyleUse::Automatic,
        &mut xml_out,
    )?;
    write_valuestyles(
        &book.formats,
        StyleOrigin::Content,
        StyleUse::Automatic,
        &mut xml_out,
    )?;
    xml_out.end_elem("office:automatic-styles")?;

    xml_out.elem("office:body")?;
    xml_out.elem("office:spreadsheet")?;

    // extra tags. pass through only
    for tag in &book.extra {
        if tag.name() == "office:scripts" ||
            tag.name() == "table:tracked-changes" ||
            tag.name() == "text:variable-decls" ||
            tag.name() == "text:sequence-decls" ||
            tag.name() == "text:user-field-decls" ||
            tag.name() == "text:dde-connection-decls" ||
            // tag.name() == "text:alphabetical-index-auto-mark-file" ||
            tag.name() == "table:calculation-settings" ||
            tag.name() == "table:label-ranges"
        {
            write_xmltag(tag, &mut xml_out)?;
        }
    }

    write_content_validations(&book, &mut xml_out)?;

    for sheet in &book.sheets {
        write_sheet(&book, sheet, &mut xml_out)?;
    }

    // extra tags. pass through only
    for tag in &book.extra {
        if tag.name() == "table:named-expressions"
            || tag.name() == "table:database-ranges"
            || tag.name() == "table:data-pilot-tables"
            || tag.name() == "table:consolidation"
            || tag.name() == "table:dde-links"
        {
            write_xmltag(tag, &mut xml_out)?;
        }
    }

    xml_out.end_elem("office:spreadsheet")?;
    xml_out.end_elem("office:body")?;
    xml_out.end_elem("office:document-content")?;

    xml_out.close()?;

    Ok(())
}

fn write_content_validations<W: Write + Seek>(
    book: &WorkBook,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if !book.validations.is_empty() {
        xml_out.elem("table:content-validations")?;

        for valid in book.validations.values() {
            xml_out.elem("table:content-validation")?;
            xml_out.attr_esc("table:name", valid.name())?;
            let mut cond = "of:".to_string();
            cond.push_str(valid.condition());
            xml_out.attr_esc("table:condition", cond.as_str())?;
            xml_out.attr(
                "table:allow-empty-cell",
                if valid.allow_empty() { "true" } else { "false" },
            )?;
            xml_out.attr(
                "table:display-list",
                match valid.display() {
                    ValidationDisplay::NoDisplay => "no",
                    ValidationDisplay::Unsorted => "unsorted",
                    ValidationDisplay::SortAscending => "sort-ascending",
                },
            )?;
            xml_out.attr_esc("table:base-cell-address", &valid.base_cell().to_string())?;

            if let Some(err) = valid.err() {
                if err.text().is_some() {
                    xml_out.elem("table:error-message")?;
                } else {
                    xml_out.empty("table:error-message")?;
                }
                xml_out.attr("table:display", err.display().to_string())?;
                xml_out.attr("table:message-type", err.msg_type().to_string())?;
                if let Some(title) = err.title() {
                    xml_out.attr("table:title", title)?;
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
                xml_out.attr("table:display", err.display().to_string())?;
                if let Some(title) = err.title() {
                    xml_out.attr("table:title", title)?;
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
fn check_hidden(ranges: &[CellRange], row: ucell, col: ucell) -> (bool, ucell) {
    if let Some(found) = ranges.iter().find(|s| s.contains(row, col)) {
        (true, found.to_col() - col)
    } else {
        (false, 0)
    }
}

/// Removes any outlived Ranges from the vector.
pub(crate) fn remove_outlooped(ranges: &mut Vec<CellRange>, row: ucell, col: ucell) {
    *ranges = ranges
        .drain(..)
        .filter(|s| !s.out_looped(row, col))
        .collect();
}

fn write_sheet<W: Write + Seek>(
    book: &WorkBook,
    sheet: &Sheet,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.elem("table:table")?;
    xml_out.attr_esc("table:name", &*sheet.name)?;
    if let Some(style) = &sheet.style {
        xml_out.attr_esc("table:style-name", style.as_str())?;
    }
    if let Some(print_ranges) = &sheet.print_ranges {
        xml_out.attr_esc("table:print-ranges", &cellranges_string(print_ranges))?;
    }
    if !sheet.print() {
        xml_out.attr("table:print", "false")?;
    }
    if !sheet.display() {
        xml_out.attr("table:display", "false")?;
    }

    let max_cell = sheet.used_grid_size();

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

    write_table_columns(&sheet, max_cell, xml_out)?;

    // list of current spans
    let mut spans = Vec::<CellRange>::new();

    // table-row + table-cell
    let mut first_cell = true;
    let mut last_r: ucell = 0;
    let mut last_r_repeat: ucell = 1;
    let mut last_c: ucell = 0;

    for ((cur_row, cur_col), cell) in sheet.data.iter() {
        // There may be a lot of gaps of any kind in our data.
        // In the XML format there is no cell identification, every gap
        // must be filled with empty rows/columns. For this we need some
        // calculations.

        // For the repeat-counter we need to look forward.
        // Works nicely with the range operator :-)
        let (next_r, next_c, is_last_cell) = if let Some(((next_r, next_c), _)) =
            sheet.data.range((*cur_row, cur_col + 1)..).next()
        {
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

        // After the first cell there is always an open row tag that
        // needs to be closed.
        if backward_dr > 0 && !first_cell {
            write_end_last_row(sheet, *cur_row, backward_dr, xml_out)?;
        }

        // Any empty rows before this one?
        if backward_dr > 0 {
            // If the last row had a repeat counter the distance is reduced.
            // We should not add any extra empty rows.
            if last_r_repeat - 1 < backward_dr {
                write_empty_rows_before(
                    sheet,
                    *cur_row,
                    first_cell,
                    backward_dr - last_r_repeat + 1,
                    max_cell,
                    xml_out,
                )?;
            }
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
        if !is_hidden && (cell.span.row_span > 1 || cell.span.col_span > 1) {
            spans.push(CellRange::origin_span(*cur_row, *cur_col, cell.span.into()));
        }

        first_cell = false;
        last_r = *cur_row;
        last_r_repeat = if let Some(row_header) = sheet.row_header.get(&cur_row) {
            row_header.repeat
        } else {
            1
        };
        last_c = *cur_col;
    }

    xml_out.end_elem("table:table")?;

    for tag in &sheet.extra {
        if tag.name() == "table:named-expressions" || tag.name() == "calcext:conditional-formats" {
            write_xmltag(tag, xml_out)?;
        }
    }

    Ok(())
}

fn write_empty_cells<W: Write + Seek>(
    mut forward_dc: u32,
    hidden_cols: u32,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
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

fn write_start_current_row<W: Write + Seek>(
    sheet: &Sheet,
    cur_row: ucell,
    backward_dc: u32,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    // Start of headers
    if let Some(header_rows) = &sheet.header_rows {
        if header_rows.row == cur_row {
            xml_out.elem("table:table-header-rows")?;
        }
    }

    xml_out.elem("table:table-row")?;
    if let Some(row_header) = sheet.row_header.get(&cur_row) {
        if row_header.repeat > 1 {
            xml_out.attr_esc("table:number-rows-repeated", &row_header.repeat.to_string())?;
        }
        if let Some(rowstyle) = row_header.style() {
            xml_out.attr_esc("table:style-name", rowstyle.as_str())?;
        }
        if let Some(cellstyle) = row_header.cellstyle() {
            xml_out.attr_esc("table:default-cell-style-name", cellstyle.as_str())?;
        }
        if row_header.visible() != Visibility::Visible {
            xml_out.attr_esc(
                "table:visibility",
                row_header.visible().to_string().as_str(),
            )?;
        }
    }

    // Might not be the first column in this row.
    if backward_dc > 0 {
        let backward_dc = backward_dc.to_string();
        xml_out.empty("table:table-cell")?;
        xml_out.attr("table:number-columns-repeated", backward_dc.as_str())?;
    }

    Ok(())
}

fn write_end_last_row<W: Write + Seek>(
    sheet: &Sheet,
    cur_row: u32,
    backward_dr: u32,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.end_elem("table:table-row")?;

    // This row was the end of the header.
    if let Some(header_rows) = &sheet.header_rows {
        let last_row = cur_row - backward_dr;
        if header_rows.to_row == last_row {
            xml_out.end_elem("table:table-header-rows")?;
        }
    }

    Ok(())
}

fn write_end_current_row<W: Write + Seek>(
    sheet: &Sheet,
    cur_row: u32,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.end_elem("table:table-row")?;

    // This row was the end of the header.
    if let Some(header_rows) = &sheet.header_rows {
        if header_rows.to_row == cur_row {
            xml_out.end_elem("table:table-header-rows")?;
        }
    }

    Ok(())
}

fn write_empty_rows_before<W: Write + Seek>(
    sheet: &Sheet,
    cur_row: ucell,
    first_cell: bool,
    mut backward_dr: u32,
    max_cell: (ucell, ucell),
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    // Empty rows in between are 1 less than the delta, except at the very start.
    let mut corr = if first_cell { 0u32 } else { 1u32 };

    // Only deltas greater 1 are relevant.
    // Or if this is the very start.
    if backward_dr > 1 || first_cell && backward_dr > 0 {
        // split up the empty rows, if there is some header stuff.
        if let Some(header_rows) = &sheet.header_rows {
            // What was the last_row? Was there a header start since?
            let last_row = cur_row - backward_dr;
            if header_rows.row < cur_row && header_rows.row > last_row {
                write_empty_row(
                    sheet,
                    last_row,
                    header_rows.row - last_row - corr,
                    max_cell,
                    xml_out,
                )?;
                xml_out.elem("table:table-header-rows")?;
                // Don't write the empty line for the first header-row, we can
                // collapse it with the rest. corr suits fine for this.
                corr = 0;
                // We correct the empty line count.
                backward_dr = cur_row - header_rows.row;
            }

            // What was the last row here? Was there a header end since?
            let last_row = cur_row - backward_dr;
            if header_rows.to_row < cur_row && header_rows.to_row > cur_row - backward_dr {
                // Empty lines, including the current line that marks
                // the end of the header.
                write_empty_row(
                    sheet,
                    last_row,
                    header_rows.to_row - last_row - corr + 1,
                    max_cell,
                    xml_out,
                )?;
                xml_out.end_elem("table:table-header-rows")?;
                // Correction for table start is no longer needed.
                corr = 1;
                // We correct the empty line count.
                backward_dr = cur_row - header_rows.to_row;
            }
        }

        // Write out the empty lines.
        let last_row = cur_row - backward_dr;
        write_empty_row(sheet, last_row, backward_dr - corr, max_cell, xml_out)?;
    }

    Ok(())
}

fn write_empty_row<W: Write + Seek>(
    sheet: &Sheet,
    cur_row: ucell,
    empty_count: u32,
    max_cell: (u32, u32),
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    xml_out.elem("table:table-row")?;
    xml_out.attr("table:number-rows-repeated", &empty_count.to_string())?;
    if let Some(row_header) = sheet.row_header.get(&cur_row) {
        if let Some(rowstyle) = row_header.style() {
            xml_out.attr_esc("table:style-name", rowstyle.as_str())?;
        }
        if let Some(cellstyle) = row_header.cellstyle() {
            xml_out.attr_esc("table:default-cell-style-name", cellstyle.as_str())?;
        }
        if row_header.visible() != Visibility::Visible {
            xml_out.attr_esc(
                "table:visibility",
                row_header.visible().to_string().as_str(),
            )?;
        }
    }

    // We fill the empty spaces completely up to max columns.
    let max_cell_col = max_cell.1.to_string();
    xml_out.empty("table:table-cell")?;
    xml_out.attr("table:number-columns-repeated", max_cell_col.as_str())?;

    xml_out.end_elem("table:table-row")?;

    Ok(())
}

fn write_xmltag<W: Write + Seek>(
    x: &XmlTag,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if x.is_empty() {
        xml_out.empty(x.name())?;
    } else {
        xml_out.elem(x.name())?;
    }
    for (k, v) in x.attrmap().iter() {
        xml_out.attr_esc(k.as_ref(), v.as_str())?;
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

fn write_table_columns<W: Write + Seek>(
    sheet: &Sheet,
    max_cell: (ucell, ucell),
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    // table:table-column
    for c in 0..max_cell.1 {
        // markup header columns
        if let Some(header_cols) = &sheet.header_cols {
            if header_cols.col() == c {
                xml_out.elem("table:table-header-columns")?;
            }
        }

        xml_out.empty("table:table-column")?;
        if let Some(col_header) = sheet.col_header.get(&c) {
            if let Some(style) = col_header.style() {
                xml_out.attr_esc("table:style-name", style.as_str())?;
            }
            if let Some(cellstyle) = col_header.cellstyle() {
                xml_out.attr_esc("table:default-cell-style-name", cellstyle.as_str())?;
            }
            if col_header.visible() != Visibility::Visible {
                xml_out.attr_esc(
                    "table:visibility",
                    col_header.visible().to_string().as_str(),
                )?;
            }
        }

        // markup header columns
        if let Some(header_cols) = &sheet.header_cols {
            if header_cols.to_col() == c {
                xml_out.end_elem("table:table-header-columns")?;
            }
        }
    }

    Ok(())
}

#[allow(clippy::single_char_add_str)]
fn write_cell<W: Write + Seek>(
    book: &WorkBook,
    cell: &CellData,
    is_hidden: bool,
    xml_out: &mut XmlOdsWriter<W>,
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

    if let Some(formula) = &cell.formula {
        xml_out.attr_esc("table:formula", formula.as_str())?;
    }

    // Direct style oder value based default style.
    if let Some(style) = &cell.style {
        xml_out.attr_esc("table:style-name", style.as_str())?;
    } else if let Some(style) = book.def_style(cell.value.value_type()) {
        xml_out.attr_esc("table:style-name", style.as_str())?;
    }
    if let Some(validation_name) = &cell.validation_name {
        xml_out.attr("table:content-validation-name", validation_name.as_str())?;
    }

    // Spans
    if cell.span.row_span > 1 {
        xml_out.attr_esc(
            "table:number-rows-spanned",
            cell.span.row_span.to_string().as_str(),
        )?;
    }
    if cell.span.col_span > 1 {
        xml_out.attr_esc(
            "table:number-columns-spanned",
            cell.span.col_span.to_string().as_str(),
        )?;
    }

    // Might not yield a useful result. Could not exist, or be in styles.xml
    // which I don't read. Defaulting to to_string() seems reasonable.
    let valuestyle = if let Some(style_name) = &cell.style {
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
        Value::TextXml(t) => {
            xml_out.attr("office:value-type", "string")?;
            for tt in t.iter() {
                write_xmltag(tt, xml_out)?;
            }
        }
        Value::DateTime(d) => {
            xml_out.attr("office:value-type", "date")?;
            let value = d.format("%Y-%m-%dT%H:%M:%S%.f").to_string();
            xml_out.attr("office:date-value", value.as_str())?;
            xml_out.elem("text:p")?;
            if let Some(valuestyle) = valuestyle {
                xml_out.text_esc(valuestyle.format_datetime(d).as_str())?;
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
            if let Some(valuestyle) = valuestyle {
                xml_out.text_esc(valuestyle.format_time_duration(d).as_str())?;
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
            if let Some(valuestyle) = valuestyle {
                xml_out.text_esc(valuestyle.format_boolean(*b).as_str())?;
            } else {
                xml_out.text(if *b { "true" } else { "false" })?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::Currency(v, c) => {
            xml_out.attr("office:value-type", "currency")?;
            xml_out.attr_esc("office:currency", String::from_utf8_lossy(c))?;
            let value = v.to_string();
            xml_out.attr("office:value", value.as_str())?;
            xml_out.elem("text:p")?;
            if let Some(valuestyle) = valuestyle {
                xml_out.text_esc(valuestyle.format_float(*v).as_str())?;
            } else {
                xml_out.text(String::from_utf8_lossy(c))?;
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
            if let Some(valuestyle) = valuestyle {
                xml_out.text_esc(valuestyle.format_float(*v).as_str())?;
            } else {
                xml_out.text(value.as_str())?;
            }
            xml_out.end_elem("text:p")?;
        }
        Value::Percentage(v) => {
            xml_out.attr("office:value-type", "percentage")?;
            xml_out.attr("office:value", format!("{}%", v).as_str())?;
            xml_out.elem("text:p")?;
            if let Some(valuestyle) = valuestyle {
                xml_out.text_esc(valuestyle.format_float(*v * 100.0).as_str())?;
            } else {
                xml_out.text(&(v * 100.0).to_string())?;
            }
            xml_out.end_elem("text:p")?;
        }
    }

    match cell.value {
        Value::Empty => {}
        _ => xml_out.end_elem(tag)?,
    }

    Ok(())
}

fn write_font_decl<W: Write + Seek>(
    fonts: &HashMap<String, FontFaceDecl>,
    origin: StyleOrigin,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    for font in fonts.values().filter(|s| s.origin() == origin) {
        xml_out.empty("style:font-face")?;
        xml_out.attr_esc("style:name", font.name().as_str())?;
        for (a, v) in font.attrmap().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    Ok(())
}

fn write_styles<W: Write + Seek>(
    book: &WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    for style in book.tablestyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_tablestyle(style, xml_out)?;
        }
    }
    for style in book.rowstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_rowstyle(style, xml_out)?;
        }
    }
    for style in book.colstyles.values() {
        if style.origin() == origin && style.styleuse() == styleuse {
            write_colstyle(style, xml_out)?;
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

fn write_tablestyle<W: Write + Seek>(
    style: &TableStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "table")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.tablestyle().is_empty() {
        xml_out.empty("style:table-properties")?;
        for (a, v) in style.tablestyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_rowstyle<W: Write + Seek>(
    style: &RowStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "table-row")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.rowstyle().is_empty() {
        xml_out.empty("style:table-row-properties")?;
        for (a, v) in style.rowstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_colstyle<W: Write + Seek>(
    style: &ColStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "table-column")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.colstyle().is_empty() {
        xml_out.empty("style:table-column-properties")?;
        for (a, v) in style.colstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_cellstyle<W: Write + Seek>(
    style: &CellStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "table-cell")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.cellstyle().is_empty() {
        xml_out.empty("style:table-cell-properties")?;
        for (a, v) in style.cellstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if !&style.paragraphstyle().is_empty() {
        xml_out.empty("style:paragraph-properties")?;
        for (a, v) in style.paragraphstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if !style.textstyle().is_empty() {
        xml_out.empty("style:text-properties")?;
        for (a, v) in style.textstyle().iter() {
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
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_paragraphstyle<W: Write + Seek>(
    style: &ParagraphStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "paragraph")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.paragraphstyle().is_empty() {
        if style.tabstops().is_none() {
            xml_out.empty("style:paragraph-properties")?;
            for (a, v) in style.paragraphstyle().iter() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        } else {
            xml_out.elem("style:paragraph-properties")?;
            for (a, v) in style.paragraphstyle().iter() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
            xml_out.elem("style:tab-stops")?;
            if let Some(tabstops) = style.tabstops() {
                for ts in tabstops {
                    xml_out.empty("style:tab-stop")?;
                    for (a, v) in ts.attrmap().iter() {
                        xml_out.attr_esc(a.as_ref(), v.as_str())?;
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
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_textstyle<W: Write + Seek>(
    style: &TextStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "text")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.textstyle().is_empty() {
        xml_out.empty("style:text-properties")?;
        for (a, v) in style.textstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }
    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_graphicstyle<W: Write + Seek>(
    style: &GraphicStyle,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if style.styleuse() == StyleUse::Default {
        xml_out.elem("style:default-style")?;
    } else {
        xml_out.elem("style:style")?;
        xml_out.attr_esc("style:name", style.name())?;
    }
    xml_out.attr("style:family", "graphic")?;
    for (a, v) in style.attrmap().iter() {
        match a.as_ref() {
            "style:name" => {}
            "style:family" => {}
            _ => {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }
    }

    if !style.graphicstyle().is_empty() {
        xml_out.empty("style:graphic-properties")?;
        for (a, v) in style.graphicstyle().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }
    }

    if style.styleuse() == StyleUse::Default {
        xml_out.end_elem("style:default-style")?;
    } else {
        xml_out.end_elem("style:style")?;
    }

    Ok(())
}

fn write_valuestyles<W: Write + Seek>(
    value_formats: &HashMap<String, ValueFormat>,
    origin: StyleOrigin,
    styleuse: StyleUse,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    for value_format in value_formats
        .values()
        .filter(|s| s.origin() == origin && s.styleuse() == styleuse)
    {
        let tag = match value_format.value_type() {
            ValueType::Empty => "number:empty_style", // ???
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
        xml_out.attr_esc("style:name", value_format.name().as_str())?;
        for (a, v) in value_format.attrmap().iter() {
            xml_out.attr_esc(a.as_ref(), v.as_str())?;
        }

        if !value_format.textstyle().is_empty() {
            xml_out.empty("style:text-properties")?;
            for (a, v) in value_format.textstyle().iter() {
                xml_out.attr_esc(a.as_ref(), v.as_str())?;
            }
        }

        for part in value_format.parts() {
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
            };

            if part.part_type() == FormatPartType::Text
                || part.part_type() == FormatPartType::CurrencySymbol
            {
                xml_out.elem(part_tag)?;
                for (a, v) in part.attrmap().iter() {
                    xml_out.attr_esc(a.as_ref(), v.as_str())?;
                }
                if let Some(content) = part.content() {
                    xml_out.text_esc(content)?;
                }
                xml_out.end_elem(part_tag)?;
            } else {
                xml_out.empty(part_tag)?;
                for (a, v) in part.attrmap().iter() {
                    xml_out.attr_esc(a.as_ref(), v.as_str())?;
                }
            }
        }

        if let Some(stylemaps) = value_format.stylemaps() {
            for sm in stylemaps {
                xml_out.empty("style:map")?;
                xml_out.attr_esc("style:condition", sm.condition())?;
                xml_out.attr_esc("style:apply-style-name", sm.applied_style())?;
                xml_out.attr_esc("style:base-cell-address", &sm.base_cell().to_string())?;
            }
        }

        xml_out.end_elem(tag)?;
    }

    Ok(())
}

fn write_pagestyles<W: Write + Seek>(
    styles: &HashMap<String, PageStyle>,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    for style in styles.values() {
        xml_out.elem("style:page-layout")?;
        xml_out.attr_esc("style:name", &style.name())?;

        if !style.style().is_empty() {
            xml_out.empty("style:page-layout-properties")?;
            for (k, v) in style.style().iter() {
                xml_out.attr(k.as_ref(), v.as_str())?;
            }
        }

        xml_out.elem("style:header-style")?;
        xml_out.empty("style:header-footer-properties")?;
        if !style.headerstyle().style().is_empty() {
            for (k, v) in style.headerstyle().style().iter() {
                xml_out.attr(k.as_ref(), v.as_str())?;
            }
        }
        xml_out.end_elem("style:header-style")?;

        xml_out.elem("style:footer-style")?;
        xml_out.empty("style:header-footer-properties")?;
        if !style.footerstyle().style().is_empty() {
            for (k, v) in style.footerstyle().style().iter() {
                xml_out.attr(k.as_ref(), v.as_str())?;
            }
        }
        xml_out.end_elem("style:footer-style")?;

        xml_out.end_elem("style:page-layout")?;
    }

    Ok(())
}

fn write_masterpage<W: Write + Seek>(
    styles: &HashMap<String, MasterPage>,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    for style in styles.values() {
        xml_out.elem("style:master-page")?;
        xml_out.attr("style:name", &style.name())?;
        xml_out.attr("style:page-layout-name", &style.pagestyle())?;

        xml_out.elem("style:header")?;
        if !style.header().display() {
            xml_out.attr("style:display", "false")?;
        }
        write_regions(&style.header(), xml_out)?;
        xml_out.end_elem("style:header")?;

        if !style.header_first().is_empty() {
            xml_out.elem("style:header_first")?;
            if !style.header_first().display() {
                xml_out.attr("style:display", "false")?;
            }
            write_regions(&style.header_first(), xml_out)?;
            xml_out.end_elem("style:header_first")?;
        }

        xml_out.elem("style:header_left")?;
        if !style.header_left().display() || style.header_left().is_empty() {
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

        if !style.footer_first().is_empty() {
            xml_out.elem("style:footer_first")?;
            if !style.footer_first().display() {
                xml_out.attr("style:display", "false")?;
            }
            write_regions(&style.footer_first(), xml_out)?;
            xml_out.end_elem("style:footer_first")?;
        }

        xml_out.elem("style:footer_left")?;
        if !style.footer_left().display() || style.footer_left().is_empty() {
            xml_out.attr("style:display", "false")?;
        }
        write_regions(&style.footer_left(), xml_out)?;
        xml_out.end_elem("style:footer_left")?;

        xml_out.end_elem("style:master-page")?;
    }

    Ok(())
}

fn write_regions<W: Write + Seek>(
    hf: &HeaderFooter,
    xml_out: &mut XmlOdsWriter<W>,
) -> Result<(), OdsError> {
    if let Some(left) = hf.left() {
        xml_out.elem("style:region-left")?;
        write_xmltag(left, xml_out)?;
        xml_out.end_elem("style:region-left")?;
    }
    if let Some(center) = hf.center() {
        xml_out.elem("style:region-center")?;
        write_xmltag(center, xml_out)?;
        xml_out.end_elem("style:region-center")?;
    }
    if let Some(right) = hf.right() {
        xml_out.elem("style:region-right")?;
        write_xmltag(right, xml_out)?;
        xml_out.end_elem("style:region-right")?;
    }
    if let Some(content) = hf.content() {
        write_xmltag(content, xml_out)?;
    }

    Ok(())
}
