use std::convert::{TryFrom, TryInto};
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, Write};
use std::path::Path;

use chrono::{Duration, NaiveDateTime};
use quick_xml::events::{BytesStart, Event};
use zip::ZipArchive;

use crate::attrmap2::AttrMap2;
use crate::condition::{Condition, ValueCondition};
use crate::config::{Config, ConfigItem, ConfigItemType, ConfigValue};
use crate::draw::Annotation;
use crate::ds::bufstack::BufStack;
use crate::ds::detach::Detach;
use crate::error::OdsError;
use crate::format::{FormatPart, FormatPartType, ValueFormatTrait, ValueStyleMap};
use crate::io::parse::{
    parse_bool, parse_currency, parse_datetime, parse_duration, parse_f64, parse_i16, parse_i32,
    parse_i64, parse_string, parse_u32, parse_visibility, parse_xlink_actuate, parse_xlink_show,
    parse_xlink_type,
};
use crate::io::{DUMP_UNUSED, DUMP_XML};
use crate::manifest::Manifest;
use crate::metadata::{
    MetaAutoReload, MetaDocumentStatistics, MetaHyperlinkBehaviour, MetaTemplate, MetaUserDefined,
    MetaValue,
};
use crate::refs::{parse_cellranges, parse_cellref};
use crate::style::stylemap::StyleMap;
use crate::style::tabstop::TabStop;
use crate::style::{
    ColStyle, FontFaceDecl, GraphicStyle, HeaderFooter, MasterPage, PageStyle, ParagraphStyle,
    RowStyle, RubyStyle, StyleOrigin, StyleUse, TableStyle, TextStyle,
};
use crate::text::{TextP, TextTag};
use crate::validation::{MessageType, Validation, ValidationError, ValidationHelp};
use crate::xmltree::XmlTag;
use crate::{
    CellData, CellStyle, EventListener, Grouped, Length, NamespaceMap, Script, Sheet, SplitMode,
    Value, ValueFormatBoolean, ValueFormatCurrency, ValueFormatDateTime, ValueFormatNumber,
    ValueFormatPercentage, ValueFormatText, ValueFormatTimeDuration, ValueType, Visibility,
    WorkBook,
};
use quick_xml::events::attributes::Attribute;
use std::borrow::Cow;
use std::str::from_utf8;
use zip::read::ZipFile;

type OdsXmlReader<'a> = quick_xml::Reader<BufReader<Box<dyn Read + 'a>>>;

/// Reads an ODS-file from a buffer
pub fn read_ods_buf(buf: &[u8]) -> Result<WorkBook, OdsError> {
    let zip = ZipArchive::new(Cursor::new(buf))?;
    read_ods_impl(zip)
}

/// Reads an ODS-file from a reader
pub fn read_ods_from<T: Read + Seek>(read: T) -> Result<WorkBook, OdsError> {
    let zip = ZipArchive::new(read)?;
    read_ods_impl(zip)
}

/// Reads an ODS-file.
pub fn read_ods<P: AsRef<Path>>(path: P) -> Result<WorkBook, OdsError> {
    let file = File::open(path.as_ref())?;
    let zip = ZipArchive::new(file)?;
    read_ods_impl(zip)
}

/// Reads an FODS-file from a buffer
#[allow(dead_code)]
pub fn read_fods_buf(buf: &[u8]) -> Result<WorkBook, OdsError> {
    let read = Cursor::new(buf);
    read_fods_impl(Box::new(read))
}

/// Reads an FODS-file from a reader
#[allow(dead_code)]
pub fn read_fods_from<T: Read>(read: T) -> Result<WorkBook, OdsError> {
    read_fods_impl(Box::new(read))
}

/// Reads an FODS-file.
#[allow(dead_code)]
pub fn read_fods<P: AsRef<Path>>(path: P) -> Result<WorkBook, OdsError> {
    let file = File::open(path.as_ref())?;
    read_fods_impl(Box::new(file))
}

/// Reads an ODS-file.
fn read_fods_impl(read: Box<dyn Read + '_>) -> Result<WorkBook, OdsError> {
    let mut book = WorkBook::new_empty();
    let mut bufstack = BufStack::new();

    let read = BufReader::new(read);
    let mut xml = quick_xml::Reader::from_reader(read);

    read_fods_content(&mut bufstack, &mut book, &mut xml)?;

    Ok(book)
}

fn read_fods_content(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();

    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!("read_fods_content {:?}", evt);
        }

        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:document" => {
                let (version, xmlns) = read_namespaces_and_version(xml_tag)?;
                book.xmlns.insert("fods.xml".to_string(), xmlns);
                if let Some(version) = version {
                    book.set_version(version);
                }
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:document" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:meta" => {
                read_office_meta(bs, book, xml)?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:settings" => {
                read_office_settings(bs, book, xml)?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:scripts" => {
                read_scripts(bs, book, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:font-face-decls" => {
                read_office_font_face_decls(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:styles" => {
                read_office_styles(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:automatic-styles" => {
                read_office_automatic_styles(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:master-styles" => {
                read_office_master_styles(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:body" => {
                read_office_body(bs, book, xml)?;
            }

            Event::Decl(_) => {}
            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_fods_content", &evt)?;
            }
        }
    }

    bs.push(buf);
    Ok(())
}

/// Reads an ODS-file.
fn read_ods_impl<R: Read + Seek>(mut zip: ZipArchive<R>) -> Result<WorkBook, OdsError> {
    let mut book = WorkBook::new_empty();
    let mut bufstack = BufStack::new();

    read_ods_manifest(&mut bufstack, &mut book, &mut zip)?;

    if let Ok(z) = zip.by_name("meta.xml") {
        let read: Box<dyn Read> = Box::new(z);
        let mut xml = quick_xml::Reader::from_reader(BufReader::new(read));
        read_ods_metadata(&mut bufstack, &mut book, &mut xml)?;
    } else {
        book.metadata = Default::default();
    }

    if let Ok(z) = zip.by_name("settings.xml") {
        let read: Box<dyn Read> = Box::new(z);
        let mut xml = quick_xml::Reader::from_reader(BufReader::new(read));
        read_ods_settings(&mut bufstack, &mut book, &mut xml)?;
    } else {
        book.config = default_settings();
    }

    {
        let read: Box<dyn Read> = Box::new(zip.by_name("styles.xml")?);
        let mut xml = quick_xml::Reader::from_reader(BufReader::new(read));
        read_ods_styles(&mut bufstack, &mut book, &mut xml)?;
    }

    {
        let read: Box<dyn Read> = Box::new(zip.by_name("content.xml")?);
        let mut xml = quick_xml::Reader::from_reader(BufReader::new(read));
        read_ods_content(&mut bufstack, &mut book, &mut xml)?;
    }

    // We do some data duplication here, to make everything easier to use.
    calc_derived(&mut book)?;

    Ok(book)
}

fn read_ods_manifest<R: Read + Seek>(
    bs: &mut BufStack,
    book: &mut WorkBook,
    zip: &mut ZipArchive<R>,
) -> Result<(), OdsError> {
    {
        let zip_file = zip.by_name("META-INF/manifest.xml")?;
        let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
        read_manifest_manifest(bs, book, &mut xml)?;
    }

    // now the data if needed ...
    for manifest in book.manifest.values_mut() {
        if !matches!(
            manifest.full_path.as_str(),
            "/" | "settings.xml" | "styles.xml" | "content.xml" | "meta.xml"
        ) && !manifest.is_dir()
        {
            let mut ze = zip.by_name(manifest.full_path.as_str())?;
            let mut buf = Vec::new();
            ze.read_to_end(&mut buf)?;
            manifest.buffer = Some(buf);
        }
    }

    Ok(())
}

fn read_manifest_manifest(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut quick_xml::Reader<BufReader<ZipFile<'_>>>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        match &evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"manifest:manifest" => {}
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"manifest:manifest" => {}

            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"manifest:file-entry" => {
                let mut manifest = Manifest::default();

                for attr in xml_tag.attributes().with_checks(false) {
                    let attr = attr?;

                    if attr.key.as_ref() == b"manifest:full-path" {
                        manifest.full_path = attr.unescape_value()?.to_string();
                    } else if attr.key.as_ref() == b"manifest:version" {
                        manifest.version = Some(attr.unescape_value()?.to_string());
                    } else if attr.key.as_ref() == b"manifest:media-type" {
                        manifest.media_type = attr.unescape_value()?.to_string();
                    }
                }

                book.add_manifest(manifest);
            }

            Event::Eof => {
                break;
            }

            _ => {
                dump_unused2("read_manifest", &evt)?;
            }
        }
        buf.clear();
    }
    bs.push(buf);
    Ok(())
}

// Sets some values from the styles on the corresponding data fields.
fn calc_derived(book: &mut WorkBook) -> Result<(), OdsError> {
    let v = book
        .config
        .get_value(&["ooo:view-settings", "Views", "0", "ActiveTable"]);
    if let Some(ConfigValue::String(n)) = v {
        book.config_mut().active_table = n.clone();
    }
    let v = book
        .config
        .get_value(&["ooo:view-settings", "Views", "0", "HasSheetTabs"]);
    if let Some(ConfigValue::Boolean(n)) = v {
        book.config_mut().has_sheet_tabs = *n;
    }
    let v = book
        .config
        .get_value(&["ooo:view-settings", "Views", "0", "ShowGrid"]);
    if let Some(ConfigValue::Boolean(n)) = v {
        book.config_mut().show_grid = *n;
    }
    let v = book
        .config
        .get_value(&["ooo:view-settings", "Views", "0", "ShowPageBreaks"]);
    if let Some(ConfigValue::Boolean(n)) = v {
        book.config_mut().show_page_breaks = *n;
    }

    for i in 0..book.num_sheets() {
        let mut sheet = book.detach_sheet(i);

        // Set the column widths.
        for ch in sheet.col_header.values_mut() {
            if let Some(style_name) = &ch.style {
                if let Some(style) = book.colstyle(style_name) {
                    if style.use_optimal_col_width()? {
                        ch.set_width(Length::Default);
                    } else {
                        ch.set_width(style.col_width()?);
                    }
                }
            }
        }

        // Set the row heights
        for rh in sheet.row_header.values_mut() {
            if let Some(style_name) = &rh.style {
                if let Some(style) = book.rowstyle(style_name) {
                    if style.use_optimal_row_height()? {
                        rh.set_height(Length::Default);
                    } else {
                        rh.set_height(style.row_height()?);
                    }
                }
            }
        }

        let v = book.config.get(&[
            "ooo:view-settings",
            "Views",
            "0",
            "Tables",
            sheet.name().as_str(),
        ]);

        if let Some(cc) = v {
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["CursorPositionX"]) {
                sheet.config_mut().cursor_x = *n as u32;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["CursorPositionY"]) {
                sheet.config_mut().cursor_y = *n as u32;
            }
            if let Some(ConfigValue::Short(n)) = cc.get_value_rec(&["HorizontalSplitMode"]) {
                sheet.config_mut().hor_split_mode = SplitMode::try_from(*n)?;
            }
            if let Some(ConfigValue::Short(n)) = cc.get_value_rec(&["VerticalSplitMode"]) {
                sheet.config_mut().vert_split_mode = SplitMode::try_from(*n)?;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["HorizontalSplitPosition"]) {
                sheet.config_mut().hor_split_pos = *n as u32;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["VerticalSplitPosition"]) {
                sheet.config_mut().vert_split_pos = *n as u32;
            }
            if let Some(ConfigValue::Short(n)) = cc.get_value_rec(&["ActiveSplitRange"]) {
                sheet.config_mut().active_split_range = *n;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["PositionLeft"]) {
                sheet.config_mut().position_left = *n as u32;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["PositionRight"]) {
                sheet.config_mut().position_right = *n as u32;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["PositionTop"]) {
                sheet.config_mut().position_top = *n as u32;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["PositionBottom"]) {
                sheet.config_mut().position_bottom = *n as u32;
            }
            if let Some(ConfigValue::Short(n)) = cc.get_value_rec(&["ZoomType"]) {
                sheet.config_mut().zoom_type = *n;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["ZoomValue"]) {
                sheet.config_mut().zoom_value = *n;
            }
            if let Some(ConfigValue::Boolean(n)) = cc.get_value_rec(&["ShowGrid"]) {
                sheet.config_mut().show_grid = *n;
            }
        }

        book.attach_sheet(sheet);
    }

    Ok(())
}

// Reads the content.xml
fn read_ods_content(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_ods_content {:?}", evt);
        }
        match &evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:document-content" => {
                let (version, xmlns) = read_namespaces_and_version(xml_tag)?;
                if let Some(version) = version {
                    book.set_version(version);
                }
                book.xmlns.insert("content.xml".to_string(), xmlns);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:document-content" => {}

            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"office:scripts" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:scripts" => {
                read_scripts(bs, book, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:font-face-decls" => {
                read_office_font_face_decls(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:styles" => {
                read_office_styles(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:automatic-styles" => {
                read_office_automatic_styles(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:master-styles" => {
                read_office_master_styles(bs, book, StyleOrigin::Content, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:body" => {
                read_office_body(bs, book, xml)?;
            }

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_ods_content", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// Reads the content.xml
fn read_office_body(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!("read_office_body {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:body" => {}
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:body" => {
                break;
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:spreadsheet" => {}
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:spreadsheet" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:content-validations" => {
                read_validations(bs, book, xml)?
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table" => {
                read_table(bs, book, xml_tag, xml)?
            }

            // from the prelude
            Event::Empty(xml_tag) | Event::Start(xml_tag)
                if xml_tag.name().as_ref() == b"table:calculation-settings"
                    || xml_tag.name().as_ref() == b"table:label-ranges"
                    || xml_tag.name().as_ref() == b"table:tracked-changes"
                    || xml_tag.name().as_ref() == b"text:alphabetical-index-auto-mark-file"
                    || xml_tag.name().as_ref() == b"text:dde-connection-decls"
                    || xml_tag.name().as_ref() == b"text:sequence-decls"
                    || xml_tag.name().as_ref() == b"text:user-field-decls"
                    || xml_tag.name().as_ref() == b"text:variable-decls" =>
            {
                let v = read_xml(bs, xml_tag, empty_tag, xml)?;
                book.extra.push(v);
            }
            // from the epilogue
            Event::Empty(xml_tag) | Event::Start(xml_tag)
                if xml_tag.name().as_ref() == b"table:consolidation"
                    || xml_tag.name().as_ref() == b"table:data-pilot-tables"
                    || xml_tag.name().as_ref() == b"table:database-ranges"
                    || xml_tag.name().as_ref() == b"table:dde-links"
                    || xml_tag.name().as_ref() == b"table:named-expressions"
                    || xml_tag.name().as_ref() == b"calcext:conditional-formats" =>
            {
                let v = read_xml(bs, xml_tag, empty_tag, xml)?;
                book.extra.push(v);
            }
            // from the prelude
            Event::End(xml_tag)
                if xml_tag.name().as_ref() == b"table:calculation-settings"
                    || xml_tag.name().as_ref() == b"table:label-ranges"
                    || xml_tag.name().as_ref() == b"table:tracked-changes"
                    || xml_tag.name().as_ref() == b"text:alphabetical-index-auto-mark-file"
                    || xml_tag.name().as_ref() == b"text:dde-connection-decls"
                    || xml_tag.name().as_ref() == b"text:sequence-decls"
                    || xml_tag.name().as_ref() == b"text:user-field-decls"
                    || xml_tag.name().as_ref() == b"text:variable-decls" => {}
            // from the epilogue
            Event::End(xml_tag)
                if xml_tag.name().as_ref() == b"table:consolidation"
                    || xml_tag.name().as_ref() == b"table:data-pilot-tables"
                    || xml_tag.name().as_ref() == b"table:database-ranges"
                    || xml_tag.name().as_ref() == b"table:dde-links"
                    || xml_tag.name().as_ref() == b"table:named-expressions" => {}

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_office_body", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

fn read_namespaces_and_version(
    super_tag: &BytesStart<'_>,
) -> Result<(Option<String>, NamespaceMap), OdsError> {
    let mut version = None;
    let mut xmlns = NamespaceMap::new();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"office:version" => {
                version = Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref().starts_with(b"xmlns:") => {
                let k = from_utf8(attr.key.as_ref())?.to_string();
                let v = attr.unescape_value()?.to_string();
                xmlns.insert(k, v);
            }
            attr if attr.key.as_ref() == b"office:mimetype" => {
                if attr.unescape_value()? != "application/vnd.oasis.opendocument.spreadsheet" {
                    return Err(OdsError::Parse(
                        "invalid content-type",
                        Some(attr.unescape_value()?.to_string()),
                    ));
                }
            }
            attr => {
                dump_unused(
                    "read_namespaces_and_version",
                    super_tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }
    Ok((version, xmlns))
}

// Reads the table.
fn read_table(
    bs: &mut BufStack,
    book: &mut WorkBook,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut sheet = Sheet::new("");

    read_table_attr(&mut sheet, super_tag)?;

    // Position within table-columns
    let mut table_col: u32 = 0;

    // Cell position
    let mut row: u32 = 0;
    let mut col: u32 = 0;
    // Rows can be repeated. In reality only empty ones ever are.
    let mut row_repeat: u32 = 1;

    // Groups can be stacked.
    let mut col_group = Vec::new();
    let mut row_group = Vec::new();

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_table {:?}", evt);
        }
        match &evt {
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table" => {
                // TODO: Maybe find a better fix for the repeat error.
                // Reset the repeat count for the last two rows to one if it exceeds
                // some arbitrary limit.
                let mut it = sheet.row_header.iter_mut().rev();
                if let Some((_row, last)) = it.next() {
                    if last.repeat > 1000 {
                        last.repeat = 1;
                    }
                }
                if let Some((_row, last)) = it.next() {
                    if last.repeat > 1000 {
                        last.repeat = 1;
                    }
                }
                break;
            }

            // Prelude
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"table:title"
                    || xml_tag.name().as_ref() == b"table:desc"
                    || xml_tag.name().as_ref() == b"table:table-source"
                    || xml_tag.name().as_ref() == b"office:dde-source"
                    || xml_tag.name().as_ref() == b"table:scenario"
                    || xml_tag.name().as_ref() == b"office:forms"
                    || xml_tag.name().as_ref() == b"table:shapes" =>
            {
                sheet.extra.push(read_xml(bs, xml_tag, empty_tag, xml)?);
            }
            Event::End(xml_tag)
                if xml_tag.name().as_ref() == b"table:title"
                    || xml_tag.name().as_ref() == b"table:desc"
                    || xml_tag.name().as_ref() == b"table:table-source"
                    || xml_tag.name().as_ref() == b"office:dde-source"
                    || xml_tag.name().as_ref() == b"table:scenario"
                    || xml_tag.name().as_ref() == b"office:forms"
                    || xml_tag.name().as_ref() == b"table:shapes" => {}

            // Epilogue
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"table:named-expressions"
                    || xml_tag.name().as_ref() == b"calcext:conditional-formats" =>
            {
                sheet.extra.push(read_xml(bs, xml_tag, empty_tag, xml)?);
            }
            Event::End(xml_tag)
                if xml_tag.name().as_ref() == b"table:named-expressions"
                    || xml_tag.name().as_ref() == b"calcext:conditional-formats" => {}

            //
            // Table Column Data
            //
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table-column-group" => {
                let v = read_table_column_group_attr(table_col, xml_tag)?;
                col_group.push(v);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table-column-group" => {
                if let Some(mut v) = col_group.pop() {
                    v.set_to(table_col - 1);
                    sheet.group_cols.push(v);
                }
            }

            // ignore for ods
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table-header-rows" => {}
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table-header-rows" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table-columns" => {}
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table-columns" => {}

            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"table:table-column" => {
                table_col = read_table_col_attr(&mut sheet, table_col, xml_tag)?;
            }

            //
            // Table row data
            //
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table-row-group" => {
                let v = read_table_row_group_attr(row, xml_tag)?;
                row_group.push(v);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table-row-group" => {
                if let Some(mut v) = row_group.pop() {
                    v.set_to(row - 1);
                    sheet.group_rows.push(v);
                } // todo: ignore?
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table-rows" => {
                // noop
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table-rows" => {
                // noop
            }

            //
            // Table cells
            //
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:table-row" => {
                row_repeat = read_table_row_attr(&mut sheet, row, xml_tag)?;
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:table-row" => {
                row += row_repeat;
                col = 0;
                row_repeat = 1;
            }

            Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"table:table-cell"
                    || xml_tag.name().as_ref() == b"table:covered-table-cell" =>
            {
                col = read_empty_table_cell(&mut sheet, row, col, xml_tag)?;
            }
            Event::Start(xml_tag)
                if xml_tag.name().as_ref() == b"table:table-cell"
                    || xml_tag.name().as_ref() == b"table:covered-table-cell" =>
            {
                col = read_table_cell2(bs, &mut sheet, row, col, xml_tag, xml)?;
            }

            _ => {
                dump_unused2("read_table", &evt)?;
            }
        }
        buf.clear();
    }
    bs.push(buf);

    book.push_sheet(sheet);

    Ok(())
}

// Reads the table attributes.
fn read_table_attr(sheet: &mut Sheet, super_tag: &BytesStart<'_>) -> Result<(), OdsError> {
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:name" => {
                sheet.set_name(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"table:style-name" => {
                sheet.set_style(&attr.unescape_value()?.as_ref().into());
            }
            attr if attr.key.as_ref() == b"table:print" => {
                sheet.set_print(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"table:display" => {
                sheet.set_display(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"table:print-ranges" => {
                let v = attr.unescape_value()?;
                sheet.print_ranges = parse_cellranges(v.as_ref())?;
            }
            attr => {
                dump_unused("read_table_attr", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    Ok(())
}

// Reads table-row attributes. Returns the repeat-count.
fn read_table_row_attr(
    sheet: &mut Sheet,
    row: u32,
    super_tag: &BytesStart<'_>,
) -> Result<u32, OdsError> {
    let mut row_repeat: u32 = 1;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            // table:default-cell-style-name 19.615, table:visibility 19.749 and xml:id 19.914.
            attr if attr.key.as_ref() == b"table:number-rows-repeated" => {
                row_repeat = parse_u32(&attr.value)?;
                if row_repeat > 1 {
                    sheet.set_row_repeat(row, row_repeat);
                }
            }
            attr if attr.key.as_ref() == b"table:style-name" => {
                let rowstyle = Some(attr.unescape_value()?.to_string());
                if let Some(rowstyle) = rowstyle {
                    sheet.set_rowstyle(row, &rowstyle.into());
                }
            }
            attr if attr.key.as_ref() == b"table:default-cell-style-name" => {
                let row_cellstyle = Some(attr.unescape_value()?.to_string());
                if let Some(row_cellstyle) = row_cellstyle {
                    sheet.set_row_cellstyle(row, &row_cellstyle.into());
                }
            }
            attr if attr.key.as_ref() == b"table:visibility" => {
                let row_visible = parse_visibility(&attr.value)?;
                if row_visible != Visibility::Visible {
                    sheet.set_row_visible(row, row_visible);
                }
            }
            attr => {
                dump_unused("read_table_row_attr", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    Ok(row_repeat)
}

// Reads the table:table-column-group attributes.
fn read_table_column_group_attr(
    table_col: u32,
    super_tag: &BytesStart<'_>,
) -> Result<Grouped, OdsError> {
    let mut display = true;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:display" => {
                display = parse_bool(&attr.value)?;
            }
            attr => {
                dump_unused(
                    "read_table_column_group_attr",
                    super_tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    Ok(Grouped {
        from: table_col,
        to: 0,
        display,
    })
}

// Reads the table:table-row-group attributes.
fn read_table_row_group_attr(row: u32, super_tag: &BytesStart<'_>) -> Result<Grouped, OdsError> {
    let mut display = true;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:display" => {
                display = parse_bool(&attr.value)?;
            }
            attr => {
                dump_unused(
                    "read_table_row_group_attr",
                    super_tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    Ok(Grouped {
        from: row,
        to: 0,
        display,
    })
}

// Reads the table-column attributes. Creates as many copies as indicated.
fn read_table_col_attr(
    sheet: &mut Sheet,
    mut table_col: u32,
    super_tag: &BytesStart<'_>,
) -> Result<u32, OdsError> {
    let mut style = None;
    let mut cellstyle = None;
    let mut repeat: u32 = 1;
    let mut visible: Visibility = Default::default();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:number-columns-repeated" => {
                repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key.as_ref() == b"table:style-name" => {
                style = Some(attr.unescape_value()?);
            }
            attr if attr.key.as_ref() == b"table:default-cell-style-name" => {
                cellstyle = Some(attr.unescape_value()?);
            }
            attr if attr.key.as_ref() == b"table:visibility" => {
                visible = parse_visibility(&attr.value)?;
            }
            attr => {
                dump_unused("read_table_col_attr", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    while repeat > 0 {
        if let Some(style) = &style {
            sheet.set_colstyle(table_col, &style.as_ref().into());
        }
        if let Some(cellstyle) = &cellstyle {
            sheet.set_col_cellstyle(table_col, &cellstyle.as_ref().into());
        }
        sheet.set_col_visible(table_col, visible);
        table_col += 1;
        repeat -= 1;
    }

    Ok(table_col)
}

#[derive(Debug)]
#[allow(variant_size_differences)]
enum TextContent2 {
    Empty,
    Text(String),
    Xml(TextTag),
    XmlVec(Vec<TextTag>),
}

#[derive(Debug)]
struct ReadTableCell2 {
    val_type: ValueType,
    val_datetime: Option<NaiveDateTime>,
    val_duration: Option<Duration>,
    val_float: Option<f64>,
    val_bool: Option<bool>,
    val_string: Option<String>,
    val_currency: Option<String>,

    content: TextContent2,
}

fn read_table_cell2(
    bs: &mut BufStack,
    sheet: &mut Sheet,
    row: u32,
    mut col: u32,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<u32, OdsError> {
    // Current cell tag
    let tag_name = super_tag.name();

    let mut cell = CellData::default();
    let mut cell_repeat: u32 = 1;

    let mut tc = ReadTableCell2 {
        val_type: ValueType::Empty,
        val_datetime: None,
        val_duration: None,
        val_float: None,
        val_bool: None,
        val_string: None,
        val_currency: None,
        content: TextContent2::Empty,
    };

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:number-columns-repeated" => {
                cell_repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key.as_ref() == b"table:number-rows-spanned" => {
                let row_span = parse_u32(&attr.value)?;
                if row_span > 1 {
                    cell.extra_mut().span.row_span = row_span;
                }
            }
            attr if attr.key.as_ref() == b"table:number-columns-spanned" => {
                let col_span = parse_u32(&attr.value)?;
                if col_span > 1 {
                    cell.extra_mut().span.col_span = col_span;
                }
            }
            attr if attr.key.as_ref() == b"table:content-validation-name" => {
                cell.extra_mut().validation_name = Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"calcext:value-type" => {
                // not used. office:value-type seems to be good enough.
            }
            attr if attr.key.as_ref() == b"office:value-type" => {
                tc.val_type = match attr.value.as_ref() {
                    b"string" => ValueType::Text,
                    b"float" => ValueType::Number,
                    b"percentage" => ValueType::Percentage,
                    b"date" => ValueType::DateTime,
                    b"time" => ValueType::TimeDuration,
                    b"boolean" => ValueType::Boolean,
                    b"currency" => ValueType::Currency,
                    other => {
                        return Err(OdsError::Parse(
                            "Unknown cell-type {:?}",
                            Some(from_utf8(other)?.into()),
                        ))
                    }
                }
            }
            attr if attr.key.as_ref() == b"office:date-value" => {
                tc.val_datetime = Some(parse_datetime(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"office:time-value" => {
                tc.val_duration = Some(parse_duration(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"office:value" => {
                tc.val_float = Some(parse_f64(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"office:boolean-value" => {
                tc.val_bool = Some(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"office:string-value" => {
                tc.val_string = Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"office:currency" => {
                tc.val_currency = Some(parse_currency(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"table:formula" => {
                cell.formula = Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"table:style-name" => {
                cell.style = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused("read_table_cell2", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_table_cell {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"text:p" => {
                let new_txt = read_text_or_tag(bs, xml_tag, false, xml)?;
                tc.content = append_text(new_txt, tc.content);
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:annotation" => {
                let annotation = read_annotation(bs, xml_tag, xml)?;
                cell.extra_mut().annotation = Some(annotation);
            }

            Event::End(xml_tag) if xml_tag.name() == tag_name => {
                parse_value2(tc, &mut cell)?;

                while cell_repeat > 1 {
                    sheet.add_cell_data(row, col, cell.clone());
                    col += 1;
                    cell_repeat -= 1;
                }
                sheet.add_cell_data(row, col, cell);
                col += 1;

                break;
            }

            Event::Eof => {
                break;
            }

            _ => {
                dump_unused2("read_table_cell", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(col)
}

fn append_text(new_txt: TextContent2, mut content: TextContent2) -> TextContent2 {
    // There can be multiple text:p elements within the cell.
    content = match content {
        TextContent2::Empty => new_txt,
        TextContent2::Text(txt) => {
            // Have a destructured text:p from before.
            // Wrap up and create list.
            let p = TextP::new().text(txt).into_xmltag();
            let mut vec = vec![p];

            match new_txt {
                TextContent2::Empty => {}
                TextContent2::Text(txt) => {
                    let p2 = TextP::new().text(txt).into_xmltag();
                    vec.push(p2);
                }
                TextContent2::Xml(xml) => {
                    vec.push(xml);
                }
                TextContent2::XmlVec(_) => {
                    unreachable!();
                }
            }
            TextContent2::XmlVec(vec)
        }
        TextContent2::Xml(xml) => {
            let mut vec = vec![xml];
            match new_txt {
                TextContent2::Empty => {}
                TextContent2::Text(txt) => {
                    let p2 = TextP::new().text(txt).into_xmltag();
                    vec.push(p2);
                }
                TextContent2::Xml(xml) => {
                    vec.push(xml);
                }
                TextContent2::XmlVec(_) => {
                    unreachable!();
                }
            }
            TextContent2::XmlVec(vec)
        }
        TextContent2::XmlVec(mut vec) => {
            match new_txt {
                TextContent2::Empty => {}
                TextContent2::Text(txt) => {
                    let p2 = TextP::new().text(txt).into_xmltag();
                    vec.push(p2);
                }
                TextContent2::Xml(xml) => {
                    vec.push(xml);
                }
                TextContent2::XmlVec(_) => {
                    unreachable!();
                }
            }
            TextContent2::XmlVec(vec)
        }
    };

    content
}

fn parse_value2(tc: ReadTableCell2, cell: &mut CellData) -> Result<(), OdsError> {
    match tc.val_type {
        ValueType::Empty => {
            // noop
        }
        ValueType::Boolean => {
            if let Some(v) = tc.val_bool {
                cell.value = Value::Boolean(v);
            } else {
                return Err(OdsError::Parse("no boolean value", None));
            }
        }
        ValueType::Number => {
            if let Some(v) = tc.val_float {
                cell.value = Value::Number(v);
            } else {
                return Err(OdsError::Parse("no float value", None));
            }
        }
        ValueType::Percentage => {
            if let Some(v) = tc.val_float {
                cell.value = Value::Percentage(v);
            } else {
                return Err(OdsError::Parse("no float value", None));
            }
        }
        ValueType::Currency => {
            if let Some(v) = tc.val_float {
                if let Some(c) = tc.val_currency {
                    cell.value = Value::Currency(v, c);
                } else {
                    cell.value = Value::Currency(v, "".to_string());
                }
            } else {
                return Err(OdsError::Parse("no float value", None));
            }
        }
        ValueType::Text => {
            if let Some(v) = tc.val_string {
                cell.value = Value::Text(v);
            } else {
                match tc.content {
                    TextContent2::Empty => {
                        // noop
                    }
                    TextContent2::Text(txt) => {
                        cell.value = Value::Text(txt);
                    }
                    TextContent2::Xml(xml) => {
                        cell.value = Value::TextXml(vec![xml]);
                    }
                    TextContent2::XmlVec(vec) => {
                        cell.value = Value::TextXml(vec);
                    }
                }
            }
        }
        ValueType::TextXml => {
            unreachable!();
        }
        ValueType::DateTime => {
            if let Some(v) = tc.val_datetime {
                cell.value = Value::DateTime(v);
            } else {
                return Err(OdsError::Parse("no datetime value", None));
            }
        }
        ValueType::TimeDuration => {
            if let Some(v) = tc.val_duration {
                cell.value = Value::TimeDuration(v);
            } else {
                return Err(OdsError::Parse("no duration value", None));
            }
        }
    }

    Ok(())
}

/// Reads a table-cell from an empty XML tag.
/// There seems to be no data associated, but it can have a style and a formula.
/// And first of all we need the repeat count for the correct placement.
fn read_empty_table_cell(
    sheet: &mut Sheet,
    row: u32,
    mut col: u32,
    super_tag: &BytesStart<'_>,
) -> Result<u32, OdsError> {
    let mut cell = None;
    let mut cell_repeat = 1;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:number-columns-repeated" => {
                cell_repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key.as_ref() == b"table:number-rows-spanned" => {
                let row_span = parse_u32(&attr.value)?;
                if row_span > 1 {
                    cell.get_or_insert_with(CellData::default)
                        .extra_mut()
                        .span
                        .row_span = row_span;
                }
            }
            attr if attr.key.as_ref() == b"table:number-columns-spanned" => {
                let col_span = parse_u32(&attr.value)?;
                if col_span > 1 {
                    cell.get_or_insert_with(CellData::default)
                        .extra_mut()
                        .span
                        .col_span = parse_u32(&attr.value)?;
                }
            }
            attr if attr.key.as_ref() == b"table:content-validation-name" => {
                cell.get_or_insert_with(CellData::default)
                    .extra_mut()
                    .validation_name = Some(attr.unescape_value()?.to_string());
            }

            attr if attr.key.as_ref() == b"table:formula" => {
                cell.get_or_insert_with(CellData::default).formula =
                    Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"table:style-name" => {
                cell.get_or_insert_with(CellData::default).style =
                    Some(attr.unescape_value()?.to_string());
            }

            attr => {
                dump_unused("read_empty_table_cell", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    if let Some(cell) = cell {
        while cell_repeat > 1 {
            sheet.add_cell_data(row, col, cell.clone());
            col += 1;
            cell_repeat -= 1;
        }
        sheet.add_cell_data(row, col, cell);
        col += 1;
    } else {
        col += cell_repeat;
    }

    Ok(col)
}

fn read_annotation(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<Annotation, OdsError> {
    let mut annotation = Annotation::new_empty();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"office:display" => {
                annotation.set_display(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"office:name" => {
                annotation.set_name(attr.unescape_value()?);
            }
            attr => {
                let k = from_utf8(attr.key.as_ref())?;
                let v = attr.unescape_value()?.to_string();
                annotation.attrmap_mut().set_attr(k, v);
            }
        }
    }

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!("read_annotation {:?}", evt);
        }
        match &evt {
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:annotation" => {
                break;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"dc:creator" =>
            {
                annotation.set_creator(read_text(bs, xml_tag, empty_tag, parse_string, xml)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"dc:date" =>
            {
                annotation.set_date(read_text(bs, xml_tag, empty_tag, parse_datetime, xml)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"text:list"
                    || xml_tag.name().as_ref() == b"text:p" =>
            {
                annotation.push_text(read_xml(bs, xml_tag, empty_tag, xml)?);
            }

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_annotation", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(annotation)
}

fn read_scripts(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!("read_scripts {:?}", evt);
        }
        match &evt {
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:scripts" => {
                break;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"office:script" =>
            {
                book.add_script(read_script(bs, xml_tag, xml)?);
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:event-listeners" => {}
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:event-listeners" => {}

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"script:event-listener" =>
            {
                book.add_event_listener(read_event_listener(xml_tag)?);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"script:event-listener" => {}

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_scripts", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

fn read_script(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<Script, OdsError> {
    let v = read_xml(bs, super_tag, false, xml)?;
    let script: Script = Script {
        script_lang: v.get_attr("script:language").cloned().unwrap_or_default(),
        script: v.into_mixed_vec(),
    };
    Ok(script)
}

// reads the page-layout tag
fn read_event_listener(super_tag: &BytesStart<'_>) -> Result<EventListener, OdsError> {
    let mut evt = EventListener::new();
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"script:event-name" => {
                evt.event_name = attr.unescape_value()?.to_string();
            }
            attr if attr.key.as_ref() == b"script:language" => {
                evt.script_lang = attr.unescape_value()?.to_string();
            }
            attr if attr.key.as_ref() == b"script:macro-name" => {
                evt.macro_name = attr.unescape_value()?.to_string();
            }
            attr if attr.key.as_ref() == b"xlink:actuate" => {
                evt.actuate = parse_xlink_actuate(attr.unescape_value()?.as_bytes())?;
            }
            attr if attr.key.as_ref() == b"xlink:href" => {
                evt.href = attr.unescape_value()?.to_string();
            }
            attr if attr.key.as_ref() == b"xlink:type" => {
                evt.link_type = parse_xlink_type(attr.unescape_value()?.as_bytes())?;
            }
            attr => {
                dump_unused("read_event_listener", super_tag.name().as_ref(), &attr)?;
            }
        }
    }
    Ok(evt)
}

// reads a font-face
fn read_office_font_face_decls(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut font: FontFaceDecl = FontFaceDecl::new_empty();
    font.set_origin(origin);

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_fonts {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:font-face" =>
            {
                let name = copy_style_attr(font.attrmap_mut(), xml_tag)?;
                font.set_name(name);
                book.add_font(font);

                font = FontFaceDecl::new_empty();
                font.set_origin(StyleOrigin::Content);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:font-face-decls" => {
                break;
            }
            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_fonts", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// reads the page-layout tag
fn read_page_style(
    bs: &mut BufStack,
    book: &mut WorkBook,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut pl = PageStyle::new_empty();
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:name" => {
                pl.set_name(&attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"style:page-usage" => {
                pl.master_page_usage = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused("read_page_style", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    let mut headerstyle = false;
    let mut footerstyle = false;

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_page_layout {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:page-layout-properties" =>
            {
                copy_attr2(pl.style_mut(), xml_tag)?;
            }
            Event::End(end_tag) if end_tag.name().as_ref() == b"style:page-layout-properties" => {}

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:header-style" =>
            {
                headerstyle = true;
            }
            Event::End(end_tag) if end_tag.name().as_ref() == b"style:header-style" => {
                headerstyle = false;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:footer-style" =>
            {
                footerstyle = true;
            }
            Event::End(end_tag) if end_tag.name().as_ref() == b"style:footer-style" => {
                footerstyle = false;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:header-footer-properties" =>
            {
                if headerstyle {
                    copy_attr2(pl.headerstyle_mut().style_mut(), xml_tag)?;
                }
                if footerstyle {
                    copy_attr2(pl.footerstyle_mut().style_mut(), xml_tag)?;
                }
            }
            Event::End(end_tag) if end_tag.name().as_ref() == b"style:header-footer-properties" => {
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:background-image" =>
            {
                // noop for now. sets the background transparent.
            }

            Event::End(end_tag) if end_tag.name().as_ref() == b"style:page-layout" => {
                break;
            }

            Event::Text(_) => (),
            Event::Eof => break,
            _ => {
                dump_unused2("read_page_layout", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    book.add_pagestyle(pl);

    Ok(())
}

fn read_validations(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut valid = Validation::new();

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_validations {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:content-validation" => {
                read_validation(&mut valid, xml_tag)?;
                book.add_validation(valid);
                valid = Validation::new();
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"table:content-validation" => {
                read_validation(&mut valid, xml_tag)?;
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:content-validation" => {
                book.add_validation(valid);
                valid = Validation::new();
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"table:error-message" =>
            {
                read_validation_error(bs, &mut valid, xml_tag, empty_tag, xml)?;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"table:help-message" =>
            {
                read_validation_help(bs, &mut valid, xml_tag, empty_tag, xml)?;
            }

            Event::End(xml_tag) if xml_tag.name().as_ref() == b"table:content-validations" => {
                break;
            }

            Event::Text(_) => (),
            Event::Eof => break,
            _ => {
                dump_unused2("read_validations", &evt)?;
            }
        }
    }
    bs.push(buf);

    Ok(())
}

fn read_validation_help(
    bs: &mut BufStack,
    valid: &mut Validation,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut vh = ValidationHelp::new();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:display" => {
                vh.set_display(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"table:title" => {
                vh.set_title(Some(attr.unescape_value()?.to_string()));
            }
            attr => {
                dump_unused("read_validations", super_tag.name().as_ref(), &attr)?;
            }
        }
    }
    let txt = read_text_or_tag(bs, super_tag, empty_tag, xml)?;
    match txt {
        TextContent2::Empty => {}
        TextContent2::Xml(txt) => {
            vh.set_text(Some(txt));
        }
        _ => {
            return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(format!(
                "table:help-message invalid {:?}",
                txt
            ))));
        }
    }

    valid.set_help(Some(vh));
    Ok(())
}

fn read_validation_error(
    bs: &mut BufStack,
    valid: &mut Validation,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut ve = ValidationError::new();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:display" => {
                ve.set_display(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"table:message-type" => {
                let mt = match attr.value.as_ref() {
                    b"stop" => MessageType::Error,
                    b"warning" => MessageType::Warning,
                    b"information" => MessageType::Info,
                    _ => {
                        return Err(OdsError::Parse(
                            "unknown message-type",
                            Some(attr.unescape_value()?.into()),
                        ))
                    }
                };
                ve.set_msg_type(mt);
            }
            attr if attr.key.as_ref() == b"table:title" => {
                ve.set_title(Some(attr.unescape_value()?.to_string()));
            }
            attr => {
                dump_unused("read_validations", super_tag.name().as_ref(), &attr)?;
            }
        }
    }
    let txt = read_text_or_tag(bs, super_tag, empty_tag, xml)?;
    match txt {
        TextContent2::Empty => {}
        TextContent2::Xml(txt) => {
            ve.set_text(Some(txt));
        }
        _ => {
            return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(format!(
                "table:error-message invalid {:?}",
                txt
            ))));
        }
    }

    valid.set_err(Some(ve));

    Ok(())
}

fn read_validation(valid: &mut Validation, super_tag: &BytesStart<'_>) -> Result<(), OdsError> {
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"table:name" => {
                valid.set_name(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"table:condition" => {
                // split off 'of:' prefix
                let v = attr.unescape_value()?.to_string();
                valid.set_condition(Condition::new(v.split_at(3).1));
            }
            attr if attr.key.as_ref() == b"table:allow-empty-cell" => {
                valid.set_allow_empty(parse_bool(&attr.value)?);
            }
            attr if attr.key.as_ref() == b"table:base-cell-address" => {
                let v = attr.unescape_value()?;
                valid.set_base_cell(parse_cellref(&v)?);
            }
            attr if attr.key.as_ref() == b"table:display-list" => {
                valid.set_display(attr.value.as_ref().try_into()?);
            }
            attr => {
                dump_unused("read_validation", super_tag.name().as_ref(), &attr)?;
            }
        }
    }
    Ok(())
}

// read the master-styles tag
fn read_office_master_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_master_styles {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:master-page" =>
            {
                read_master_page(bs, book, origin, xml_tag, xml)?;
            }
            Event::End(e) if e.name().as_ref() == b"office:master-styles" => {
                break;
            }
            Event::Text(_) => (),
            Event::Eof => break,
            _ => {
                dump_unused2("read_master_styles", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// read the master-page tag
fn read_master_page(
    bs: &mut BufStack,
    book: &mut WorkBook,
    _origin: StyleOrigin,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut masterpage = MasterPage::new_empty();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:name" => {
                masterpage.set_name(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"style:page-layout-name" => {
                masterpage.set_pagestyle(&attr.unescape_value()?.as_ref().into());
            }
            attr if attr.key.as_ref() == b"style:display-name" => {
                masterpage.set_display_name(attr.unescape_value()?.as_ref().into());
            }
            attr => {
                dump_unused("read_master_page", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_master_page {:?}", evt);
        }
        match &evt {
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"style:header" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"style:header" => {
                masterpage.set_header(read_headerfooter(bs, xml_tag, xml)?);
            }
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"style:header-first" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"style:header-first" => {
                masterpage.set_header_first(read_headerfooter(bs, xml_tag, xml)?);
            }
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"style:header-left" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"style:header-left" => {
                masterpage.set_header_left(read_headerfooter(bs, xml_tag, xml)?);
            }
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"style:footer" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"style:footer" => {
                masterpage.set_footer(read_headerfooter(bs, xml_tag, xml)?);
            }
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"style:footer-first" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"style:footer-first" => {
                masterpage.set_footer_first(read_headerfooter(bs, xml_tag, xml)?);
            }
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"style:footer-left" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"style:footer-left" => {
                masterpage.set_footer_left(read_headerfooter(bs, xml_tag, xml)?);
            }

            Event::End(end_tag) if end_tag.name().as_ref() == b"style:master-page" => {
                break;
            }
            Event::Eof => break,
            _ => {
                dump_unused2("read_master_page", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    book.add_masterpage(masterpage);

    Ok(())
}

// reads any header or footer tags
fn read_headerfooter(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<HeaderFooter, OdsError> {
    let mut hf = HeaderFooter::new();
    let mut content = TextContent2::Empty;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:display" => {
                hf.set_display(parse_bool(&attr.value)?);
            }
            attr => {
                dump_unused("read_headerfooter", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_headerfooter {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:region-left" =>
            {
                let reg = read_xml(bs, xml_tag, empty_tag, xml)?;
                hf.set_left(reg.into_vec()?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:region-center" =>
            {
                let reg = read_xml(bs, xml_tag, empty_tag, xml)?;
                hf.set_center(reg.into_vec()?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:region-right" =>
            {
                let reg = read_xml(bs, xml_tag, empty_tag, xml)?;
                hf.set_right(reg.into_vec()?);
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"text:p" =>
            {
                let new_txt = read_text_or_tag(bs, xml_tag, empty_tag, xml)?;
                content = append_text(new_txt, content);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"text:h" =>
            {
                let new_txt = read_text_or_tag(bs, xml_tag, empty_tag, xml)?;
                content = append_text(new_txt, content);
            }
            // no other tags supported for now. they have never been seen in the wild.
            Event::Text(_) => (),
            Event::End(end_tag) => {
                if end_tag.name() == super_tag.name() {
                    hf.set_content(match content {
                        TextContent2::Empty => Vec::new(),
                        TextContent2::Text(v) => vec![TextP::new().text(v).into()],
                        TextContent2::Xml(v) => vec![v],
                        TextContent2::XmlVec(v) => v,
                    });
                    break;
                }
            }
            Event::Eof => break,
            _ => {
                dump_unused2("read_headerfooter", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(hf)
}

// reads the office-styles tag
fn read_office_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_styles_tag {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:style" =>
            {
                read_style_style(bs, book, origin, StyleUse::Named, xml_tag, empty_tag, xml)?;
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:default-style" =>
            {
                read_style_style(bs, book, origin, StyleUse::Default, xml_tag, empty_tag, xml)?;
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:boolean-style"
                    || xml_tag.name().as_ref() == b"number:date-style"
                    || xml_tag.name().as_ref() == b"number:time-style"
                    || xml_tag.name().as_ref() == b"number:number-style"
                    || xml_tag.name().as_ref() == b"number:currency-style"
                    || xml_tag.name().as_ref() == b"number:percentage-style"
                    || xml_tag.name().as_ref() == b"number:text-style" =>
            {
                read_value_format(bs, book, origin, StyleUse::Named, xml_tag, xml)?;
            }
            Event::End(e) if e.name().as_ref() == b"office:styles" => {
                break;
            }
            Event::Text(_) => (),
            Event::Eof => break,
            _ => {
                dump_unused2("read_styles_tag", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// read the automatic-styles tag
fn read_office_automatic_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_auto_styles {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:style" =>
            {
                read_style_style(
                    bs,
                    book,
                    origin,
                    StyleUse::Automatic,
                    xml_tag,
                    empty_tag,
                    xml,
                )?;
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:boolean-style"
                    || xml_tag.name().as_ref() == b"number:date-style"
                    || xml_tag.name().as_ref() == b"number:time-style"
                    || xml_tag.name().as_ref() == b"number:number-style"
                    || xml_tag.name().as_ref() == b"number:currency-style"
                    || xml_tag.name().as_ref() == b"number:percentage-style"
                    || xml_tag.name().as_ref() == b"number:text-style" =>
            {
                read_value_format(bs, book, origin, StyleUse::Automatic, xml_tag, xml)?;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:page-layout" =>
            {
                read_page_style(bs, book, xml_tag, xml)?;
            }

            Event::End(e) if e.name().as_ref() == b"office:automatic-styles" => {
                break;
            }
            Event::Text(_) => (),
            Event::Eof => break,
            _ => {
                dump_unused2("read_auto_styles", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// Reads any of the number:xxx tags
fn read_value_format(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    match super_tag.name().as_ref() {
        b"number:boolean-style" => {
            let mut valuestyle = ValueFormatBoolean::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_boolean_format(valuestyle);
        }
        b"number:date-style" => {
            let mut valuestyle = ValueFormatDateTime::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_datetime_format(valuestyle);
        }
        b"number:time-style" => {
            let mut valuestyle = ValueFormatTimeDuration::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_timeduration_format(valuestyle);
        }
        b"number:number-style" => {
            let mut valuestyle = ValueFormatNumber::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_number_format(valuestyle);
        }
        b"number:currency-style" => {
            let mut valuestyle = ValueFormatCurrency::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_currency_format(valuestyle);
        }
        b"number:percentage-style" => {
            let mut valuestyle = ValueFormatPercentage::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_percentage_format(valuestyle);
        }
        b"number:text-style" => {
            let mut valuestyle = ValueFormatText::new_empty();
            read_value_format_parts(bs, origin, styleuse, &mut valuestyle, super_tag, xml)?;
            book.add_text_format(valuestyle);
        }
        _ => {
            if DUMP_UNUSED {
                let tag_name = super_tag.name();
                let n = xml.decoder().decode(tag_name.as_ref())?;
                println!(" read_value_format unused {n}");
            }
        }
    }

    Ok(())
}

// Reads any of the number:xxx tags
fn read_value_format_parts<T: ValueFormatTrait>(
    bs: &mut BufStack,
    origin: StyleOrigin,
    styleuse: StyleUse,
    valuestyle: &mut T,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    //
    valuestyle.set_origin(origin);
    valuestyle.set_styleuse(styleuse);
    let name = copy_style_attr(valuestyle.attrmap_mut(), super_tag)?;
    valuestyle.set_name(name.as_str());

    // Styles with content information are stored before completion.
    let mut valuestyle_part = None;

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_value_format_parts {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:boolean" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Boolean)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:number" =>
            {
                // TODO: embedded-text below is completeley broken. removed it as a quick fix.
                //       embedded-text is a sub-tag of number:number. correct someday.
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Number)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:scientific-number" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::ScientificNumber)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:day" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Day)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:month" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Month)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:year" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Year)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:era" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Era)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:day-of-week" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::DayOfWeek)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:week-of-year" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::WeekOfYear)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:quarter" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Quarter)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:hours" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Hours)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:minutes" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Minutes)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:seconds" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Seconds)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:fraction" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::Fraction)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:am-pm" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::AmPm)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:text-content" =>
            {
                valuestyle.push_part(read_part(xml_tag, FormatPartType::TextContent)?);
            }
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:map" =>
            {
                valuestyle.push_stylemap(read_valuestylemap(xml_tag)?);
            }

            Event::End(xml_tag) if xml_tag.name().as_ref() == b"loext:text" => {}
            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"loext:text" => {}

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:fill-character"
                    || xml_tag.name().as_ref() == b"loext:fill-character" =>
            {
                valuestyle_part = Some(read_part(xml_tag, FormatPartType::FillCharacter)?);
                if empty_tag {
                    if let Some(part) = valuestyle_part {
                        valuestyle.push_part(part);
                    }
                    valuestyle_part = None;
                }
            }
            Event::End(e)
                if e.name().as_ref() == b"number:fill-character"
                    || e.name().as_ref() == b"loext:fill-character" =>
            {
                if let Some(part) = valuestyle_part {
                    valuestyle.push_part(part);
                }
                valuestyle_part = None;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:currency-symbol" =>
            {
                valuestyle_part = Some(read_part(xml_tag, FormatPartType::CurrencySymbol)?);
                if empty_tag {
                    if let Some(part) = valuestyle_part {
                        valuestyle.push_part(part);
                    }
                    valuestyle_part = None;
                }
            }
            Event::End(e) if e.name().as_ref() == b"number:currency-symbol" => {
                if let Some(part) = valuestyle_part {
                    valuestyle.push_part(part);
                }
                valuestyle_part = None;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"number:text" =>
            {
                valuestyle_part = Some(read_part(xml_tag, FormatPartType::Text)?);
                if empty_tag {
                    if let Some(part) = valuestyle_part {
                        valuestyle.push_part(part);
                    }
                    valuestyle_part = None;
                }
            }
            Event::End(e) if e.name().as_ref() == b"number:text" => {
                if let Some(part) = valuestyle_part {
                    valuestyle.push_part(part);
                }
                valuestyle_part = None;
            }

            Event::Start(xml_tag) | Event::Empty(xml_tag)
                if xml_tag.name().as_ref() == b"style:text-properties" =>
            {
                copy_attr2(valuestyle.textstyle_mut(), xml_tag)?;
            }

            Event::Text(e) => {
                if let Some(part) = &mut valuestyle_part {
                    part.set_content(xml.decoder().decode(e.unescape()?.as_bytes())?);
                }
            }

            Event::End(e) if e.name() == super_tag.name() => {
                break;
            }

            Event::Eof => break,
            _ => {
                dump_unused2("read_value_format_parts", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

fn read_part(
    super_tag: &BytesStart<'_>,
    part_type: FormatPartType,
) -> Result<FormatPart, OdsError> {
    let mut part = FormatPart::new(part_type);
    copy_attr2(part.attrmap_mut(), super_tag)?;

    // // There is one relevant subtag embedded-text.
    // let mut buf = bs.get_buf();
    // loop {
    //     let evt = xml.read_event_into(&mut buf)?;
    //     if DUMP_XML {
    //         println!(" read_part {:?}", evt);
    //     }
    //     match &evt {
    //         Event::Start( xml_tag2) | Event::Empty( xml_tag2) => {
    //             match xml_tag2.name().as_ref() {
    //                 b"number:embedded-text" => {
    //                     for attr in xml_tag2.attributes().with_checks(false) {
    //                         let attr = attr?;
    //                         match attr.key.as_ref() {
    //                             b"number:position" => {
    //                                 part.set_position(parse_i32(&attr.value)?);
    //                             }
    //                             attr => {
    //                                 return Err(OdsError::Ods(format!(
    //                                     "embedded-text: attribute unknown {} ",
    //                                     from_utf8(attr)?
    //                                 )))
    //                             }
    //                         }
    //                     }
    //                 }
    //                 _ => dump_unused2("read_value_format", &evt)?,
    //             }
    //         }
    //         Event::Text( e) => {
    //             part.set_content(xml.decoder().decode(e.unescape()?.as_bytes())?);
    //         }
    //         Event::End( e) => match e.name().as_ref() {
    //             b"number:embedded-text" => {
    //                 break;
    //             }
    //             _ => {
    //                 dump_unused2("read_value_format", &evt)?;
    //             }
    //         },
    //         Event::Eof => break,
    //         _ => {}
    //     }
    // }
    //
    // bs.push(buf);

    Ok(part)
}

#[allow(clippy::too_many_arguments)]
// style:style tag
fn read_style_style(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:family" => {
                match attr.value.as_ref() {
                    b"table" => {
                        read_tablestyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"table-column" => {
                        read_colstyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"table-row" => {
                        read_rowstyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"table-cell" => {
                        read_cellstyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"graphic" => {
                        read_graphicstyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"paragraph" => {
                        read_paragraphstyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"text" => {
                        read_textstyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    b"ruby" => {
                        read_rubystyle(bs, book, origin, style_use, super_tag, empty_tag, xml)?
                    }
                    value => {
                        return Err(OdsError::Ods(format!(
                            "style:family unknown {} ",
                            from_utf8(value)?
                        )))
                    }
                };
            }
            _ => {
                // not read here
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_tablestyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = TableStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_tablestyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_table_style {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag) => match xml_tag.name().as_ref() {
                    b"style:table-properties" => copy_attr2(style.tablestyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_table_style", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(end_tag) => {
                    if end_tag.name().as_ref() == super_tag.name().as_ref() {
                        book.add_tablestyle(style);
                        break;
                    } else {
                        dump_unused2("read_table_style", &evt)?;
                    }
                }
                Event::Eof => break,
                _ => {
                    dump_unused2("read_table_style", &evt)?;
                }
            }
        }

        bs.push(buf);
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_rowstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = RowStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_rowstyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_rowstyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag) => match xml_tag.name().as_ref() {
                    b"style:table-row-properties" => copy_attr2(style.rowstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_rowstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(end_tag) => {
                    if end_tag.name() == super_tag.name() {
                        book.add_rowstyle(style);
                        break;
                    } else {
                        dump_unused2("read_rowstyle", &evt)?;
                    }
                }
                Event::Eof => break,
                _ => {
                    dump_unused2("read_rowstyle", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_colstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = ColStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_colstyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_colstyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag) => match xml_tag.name().as_ref() {
                    b"style:table-column-properties" => copy_attr2(style.colstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_colstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(end_tag) => {
                    if end_tag.name() == super_tag.name() {
                        book.add_colstyle(style);
                        break;
                    } else {
                        dump_unused2("read_colstyle", &evt)?;
                    }
                }
                Event::Eof => break,
                _ => {
                    dump_unused2("read_colstyle", &evt)?;
                }
            }
        }

        bs.push(buf);
    }
    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_cellstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = CellStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_cellstyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_cellstyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:table-cell-properties" =>
                {
                    copy_attr2(style.cellstyle_mut(), xml_tag)?;
                }
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:text-properties" =>
                {
                    copy_attr2(style.textstyle_mut(), xml_tag)?;
                }
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:paragraph-properties" =>
                {
                    copy_attr2(style.paragraphstyle_mut(), xml_tag)?;
                }
                Event::End(end_tag) if end_tag.name().as_ref() == b"style:paragraph-properties" => {
                }
                // Event::Start(xml_tag) | Event::Empty(xml_tag)
                //     if xml_tag.name().as_ref() == b"style:graphic-properties" =>
                // {
                //     copy_attr(style.graphic_mut(), xml, xml_tag)?;
                // }
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:map" =>
                {
                    style.push_stylemap(read_stylemap(xml_tag)?);
                }
                // todo: tab-stops
                // b"style:tab-stops" => (),
                // b"style:tab-stop" => {
                //     let mut ts = TabStop::new();
                //     copy_attr(&mut ts, xml, xml_tag)?;
                //     style.paragraph_mut().add_tabstop(ts);
                // }
                Event::Text(_) => (),
                Event::End(end_tag) if end_tag.name() == super_tag.name() => {
                    book.add_cellstyle(style);
                    break;
                }
                Event::Eof => break,
                _ => {
                    dump_unused2("read_cellstyle", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_paragraphstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = ParagraphStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_paragraphstyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_paragraphstyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:text-properties" =>
                {
                    copy_attr2(style.textstyle_mut(), xml_tag)?;
                }
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:paragraph-properties" =>
                {
                    copy_attr2(style.paragraphstyle_mut(), xml_tag)?;
                }
                Event::End(end_tag) if end_tag.name().as_ref() == b"style:paragraph-properties" => {
                }
                // b"style:graphic-properties" => copy_attr(style.graphic_mut(), xml, xml_tag)?,
                // b"style:map" => style.push_stylemap(read_stylemap(xml, xml_tag)?),
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:tab-stops" => {}
                Event::End(end_tag) if end_tag.name().as_ref() == b"style:tab-stops" => {}
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:tab-stop" =>
                {
                    let mut ts = TabStop::new();
                    copy_attr2(ts.attrmap_mut(), xml_tag)?;
                    style.add_tabstop(ts);
                }

                Event::End(end_tag) if end_tag.name() == super_tag.name() => {
                    book.add_paragraphstyle(style);
                    break;
                }

                Event::Text(_) => (),
                Event::Eof => break,
                _ => {
                    dump_unused2("read_paragraphstyle", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_textstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = TextStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_textstyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_textstyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:text-properties" =>
                {
                    copy_attr2(style.textstyle_mut(), xml_tag)?;
                }
                Event::End(end_tag) if end_tag.name() == super_tag.name() => {
                    book.add_textstyle(style);
                    break;
                }
                Event::Text(_) => (),
                Event::Eof => break,
                _ => {
                    dump_unused2("read_textstyle", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_rubystyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = RubyStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_rubystyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_rubystyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:ruby-properties" =>
                {
                    copy_attr2(style.rubystyle_mut(), xml_tag)?;
                }
                Event::End(end_tag) if end_tag.name() == super_tag.name() => {
                    book.add_rubystyle(style);
                    break;
                }
                Event::Text(_) => (),
                Event::Eof => break,
                _ => {
                    dump_unused2("read_rubystyle", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
#[allow(clippy::too_many_arguments)]
fn read_graphicstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    style_use: StyleUse,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut style = GraphicStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(style_use);
    let name = copy_style_attr(style.attrmap_mut(), super_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_graphicstyle(style);
    } else {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_graphicstyle {:?}", evt);
            }
            match &evt {
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:graphic-properties" =>
                {
                    copy_attr2(style.graphicstyle_mut(), xml_tag)?;
                }

                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:paragraph-properties" =>
                {
                    // todo:
                }
                Event::End(xml_tag) if xml_tag.name().as_ref() == b"style:paragraph-properties" => {
                    // todo:
                }
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:text-properties" =>
                {
                    // todo:
                }
                Event::Start(xml_tag) | Event::Empty(xml_tag)
                    if xml_tag.name().as_ref() == b"style:tab-stops" =>
                {
                    // todo:
                }

                Event::End(end_tag) if end_tag.name() == super_tag.name() => {
                    book.add_graphicstyle(style);
                    break;
                }
                Event::Text(_) => (),
                Event::Eof => break,
                _ => {
                    dump_unused2("read_graphicstyle", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(())
}

fn read_valuestylemap(super_tag: &BytesStart<'_>) -> Result<ValueStyleMap, OdsError> {
    let mut sm = ValueStyleMap::default();
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:condition" => {
                sm.set_condition(ValueCondition::new(attr.unescape_value()?.to_string()));
            }
            attr if attr.key.as_ref() == b"style:apply-style-name" => {
                sm.set_applied_style(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused("read_stylemap", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    Ok(sm)
}

fn read_stylemap(super_tag: &BytesStart<'_>) -> Result<StyleMap, OdsError> {
    let mut sm = StyleMap::default();
    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:condition" => {
                sm.set_condition(ValueCondition::new(attr.unescape_value()?.to_string()));
            }
            attr if attr.key.as_ref() == b"style:apply-style-name" => {
                sm.set_applied_style(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"style:base-cell-address" => {
                let v = attr.unescape_value()?;
                sm.set_base_cell(Some(parse_cellref(v.as_ref())?));
            }
            attr => {
                dump_unused("read_stylemap", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    Ok(sm)
}

/// Copies all attributes to the map, excluding "style:name" which is returned.
fn copy_style_attr(attrmap: &mut AttrMap2, super_tag: &BytesStart<'_>) -> Result<String, OdsError> {
    let mut name = None;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"style:name" => {
                name = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                let k = from_utf8(attr.key.as_ref())?;
                let v = attr.unescape_value()?.to_string();
                attrmap.set_attr(k, v);
            }
        }
    }

    Ok(name.unwrap_or_default())
}

/// Copies all attributes to the given map.
fn copy_attr2(attrmap: &mut AttrMap2, super_tag: &BytesStart<'_>) -> Result<(), OdsError> {
    for attr in super_tag.attributes().with_checks(false) {
        let attr = attr?;

        let k = from_utf8(attr.key.as_ref())?;
        let v = attr.unescape_value()?.to_string();
        attrmap.set_attr(k, v);
    }

    Ok(())
}

fn read_ods_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_styles {:?}", evt);
        }
        match &evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:document-styles" => {
                let (_, xmlns) = read_namespaces_and_version(xml_tag)?;
                book.xmlns.insert("styles.xml".to_string(), xmlns);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:document-styles" => {
                // noop
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:font-face-decls" => {
                read_office_font_face_decls(bs, book, StyleOrigin::Styles, xml)?
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:styles" => {
                read_office_styles(bs, book, StyleOrigin::Styles, xml)?
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:automatic-styles" => {
                read_office_automatic_styles(bs, book, StyleOrigin::Styles, xml)?
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:master-styles" => {
                read_office_master_styles(bs, book, StyleOrigin::Styles, xml)?
            }

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_styles", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

#[allow(unused_variables)]
pub(crate) fn default_settings() -> Detach<Config> {
    let mut dc = Detach::new(Config::new());
    let p0 = dc.create_path(&[("ooo:view-settings", ConfigItemType::Set)]);
    p0.insert("VisibleAreaTop", 0);
    p0.insert("VisibleAreaLeft", 0);
    p0.insert("VisibleAreaWidth", 2540);
    p0.insert("VisibleAreaHeight", 1270);

    let p0 = dc.create_path(&[
        ("ooo:view-settings", ConfigItemType::Set),
        ("Views", ConfigItemType::Vec),
        ("0", ConfigItemType::Entry),
    ]);
    p0.insert("ViewId", "view1");
    let p0 = dc.create_path(&[
        ("ooo:view-settings", ConfigItemType::Set),
        ("Views", ConfigItemType::Vec),
        ("0", ConfigItemType::Entry),
        ("Tables", ConfigItemType::Map),
    ]);
    let p0 = dc.create_path(&[
        ("ooo:view-settings", ConfigItemType::Set),
        ("Views", ConfigItemType::Vec),
        ("0", ConfigItemType::Entry),
    ]);
    p0.insert("ActiveTable", "");
    p0.insert("HorizontalScrollbarWidth", 702);
    p0.insert("ZoomType", 0i16);
    p0.insert("ZoomValue", 100);
    p0.insert("PageViewZoomValue", 60);
    p0.insert("ShowPageBreakPreview", false);
    p0.insert("ShowZeroValues", true);
    p0.insert("ShowNotes", true);
    p0.insert("ShowGrid", true);
    p0.insert("GridColor", 12632256);
    p0.insert("ShowPageBreaks", false);
    p0.insert("HasColumnRowHeaders", true);
    p0.insert("HasSheetTabs", true);
    p0.insert("IsOutlineSymbolsSet", true);
    p0.insert("IsValueHighlightingEnabled", false);
    p0.insert("IsSnapToRaster", false);
    p0.insert("RasterIsVisible", false);
    p0.insert("RasterResolutionX", 1000);
    p0.insert("RasterResolutionY", 1000);
    p0.insert("RasterSubdivisionX", 1);
    p0.insert("RasterSubdivisionY", 1);
    p0.insert("IsRasterAxisSynchronized", true);
    p0.insert("AnchoredTextOverflowLegacy", false);

    let p0 = dc.create_path(&[("ooo:configuration-settings", ConfigItemType::Set)]);
    p0.insert("HasSheetTabs", true);
    p0.insert("ShowNotes", true);
    p0.insert("EmbedComplexScriptFonts", true);
    p0.insert("ShowZeroValues", true);
    p0.insert("ShowGrid", true);
    p0.insert("GridColor", 12632256);
    p0.insert("ShowPageBreaks", false);
    p0.insert("IsKernAsianPunctuation", false);
    p0.insert("LinkUpdateMode", 3i16);
    p0.insert("HasColumnRowHeaders", true);
    p0.insert("EmbedLatinScriptFonts", true);
    p0.insert("IsOutlineSymbolsSet", true);
    p0.insert("EmbedLatinScriptFonts", true);
    p0.insert("IsOutlineSymbolsSet", true);
    p0.insert("IsSnapToRaster", false);
    p0.insert("RasterIsVisible", false);
    p0.insert("RasterResolutionX", 1000);
    p0.insert("RasterResolutionY", 1000);
    p0.insert("RasterSubdivisionX", 1);
    p0.insert("RasterSubdivisionY", 1);
    p0.insert("IsRasterAxisSynchronized", true);
    p0.insert("AutoCalculate", true);
    p0.insert("ApplyUserData", true);
    p0.insert("PrinterName", "");
    p0.insert("PrinterSetup", ConfigValue::Base64Binary("".to_string()));
    p0.insert("SaveThumbnail", true);
    p0.insert("CharacterCompressionType", 0i16);
    p0.insert("SaveVersionOnClose", false);
    p0.insert("UpdateFromTemplate", true);
    p0.insert("AllowPrintJobCancel", true);
    p0.insert("LoadReadonly", false);
    p0.insert("IsDocumentShared", false);
    p0.insert("EmbedFonts", false);
    p0.insert("EmbedOnlyUsedFonts", false);
    p0.insert("EmbedAsianScriptFonts", true);
    p0.insert("SyntaxStringRef", 7i16);

    dc
}

fn read_ods_metadata(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();

    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!("read_ods_metadata {:?}", evt);
        }

        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:document-meta" => {
                let (_, xmlns) = read_namespaces_and_version(xml_tag)?;
                book.xmlns.insert("meta.xml".to_string(), xmlns);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:document-meta" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:meta" => {
                read_office_meta(bs, book, xml)?;
            }

            Event::Decl(_) => {}
            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_ods_metadata", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

fn read_office_meta(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();

    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!("read_metadata {:?}", evt);
        }

        match &evt {
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:meta" => {
                break;
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:generator" => {
                book.metadata.generator = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"dc:title" => {
                book.metadata.title = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"dc:description" => {
                book.metadata.description = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"dc:subject" => {
                book.metadata.subject = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:keyword" => {
                book.metadata.keyword = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:initial-creator" => {
                book.metadata.initial_creator = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"dc:creator" => {
                book.metadata.creator = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:printed-by" => {
                book.metadata.printed_by = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:creation-date" => {
                book.metadata.creation_date = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(Some(parse_datetime(v)?)),
                    || None,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"dc:date" => {
                book.metadata.date = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(Some(parse_datetime(v)?)),
                    || None,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:print-date" => {
                book.metadata.print_date = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(Some(parse_datetime(v)?)),
                    || None,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"dc:language" => {
                book.metadata.language = read_metadata_value(
                    bs,
                    xml_tag,
                    xml,
                    |v| Ok(from_utf8(v)?.to_string()),
                    String::new,
                )?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:editing-cycles" => {
                book.metadata.editing_cycles =
                    read_metadata_value(bs, xml_tag, xml, parse_u32, || 0)?;
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:editing-duration" => {
                book.metadata.editing_duration =
                    read_metadata_value(bs, xml_tag, xml, parse_duration, || Duration::seconds(0))?;
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:template" => {
                book.metadata.template = read_metadata_template(xml_tag)?;
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"meta:template" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:auto-reload" => {
                book.metadata.auto_reload = read_metadata_auto_reload(xml_tag)?;
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"meta:auto-reload" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:hyperlink-behaviour" => {
                book.metadata.hyperlink_behaviour = read_metadata_hyperlink_behaviour(xml_tag)?;
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"meta:hyperlink-behaviour" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:document-statistic" => {
                book.metadata.document_statistics = read_metadata_document_statistics(xml_tag)?;
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"meta:document-statistic" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"meta:user-defined" => {
                book.metadata
                    .user_defined
                    .push(read_metadata_user_defined(bs, xml_tag, xml)?);
            }

            Event::Empty(_) => {}

            Event::Text(_) => {}

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_metadata", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

fn read_metadata_template(tag: &BytesStart<'_>) -> Result<MetaTemplate, OdsError> {
    let mut template = MetaTemplate::default();

    for attr in tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"meta:date" => {
                dbg!(&attr);
                template.date = Some(parse_datetime(attr.unescape_value()?.as_bytes())?);
            }
            attr if attr.key.as_ref() == b"xlink:actuate" => {
                template.actuate = Some(parse_xlink_actuate(attr.unescape_value()?.as_bytes())?);
            }
            attr if attr.key.as_ref() == b"xlink:href" => {
                template.href = Some(attr.unescape_value()?.to_string())
            }
            attr if attr.key.as_ref() == b"xlink:title" => {
                template.title = Some(attr.unescape_value()?.to_string())
            }
            attr if attr.key.as_ref() == b"xlink:type" => {
                template.link_type = Some(parse_xlink_type(attr.unescape_value()?.as_bytes())?);
            }
            attr => {
                dump_unused("read_metadata_template", tag.name().as_ref(), &attr)?;
            }
        }
    }

    Ok(template)
}

fn read_metadata_auto_reload(tag: &BytesStart<'_>) -> Result<MetaAutoReload, OdsError> {
    let mut auto_reload = MetaAutoReload::default();

    for attr in tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"meta:delay" => {
                auto_reload.delay = Some(parse_duration(attr.unescape_value()?.as_bytes())?);
            }
            attr if attr.key.as_ref() == b"xlink:actuate" => {
                auto_reload.actuate = Some(parse_xlink_actuate(attr.unescape_value()?.as_bytes())?);
            }
            attr if attr.key.as_ref() == b"xlink:href" => {
                auto_reload.href = Some(attr.unescape_value()?.to_string())
            }
            attr if attr.key.as_ref() == b"xlink:show" => {
                auto_reload.show = Some(parse_xlink_show(attr.unescape_value()?.as_bytes())?);
            }
            attr if attr.key.as_ref() == b"xlink:type" => {
                auto_reload.link_type = Some(parse_xlink_type(attr.unescape_value()?.as_bytes())?);
            }
            attr => {
                dump_unused("read_metadata_auto_reload", tag.name().as_ref(), &attr)?;
            }
        }
    }

    Ok(auto_reload)
}

fn read_metadata_hyperlink_behaviour(
    tag: &BytesStart<'_>,
) -> Result<MetaHyperlinkBehaviour, OdsError> {
    let mut hyperlink_behaviour = MetaHyperlinkBehaviour::default();

    for attr in tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"office:targetframe-name" => {
                hyperlink_behaviour.target_frame_name = Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"xlink:show" => {
                hyperlink_behaviour.show =
                    Some(parse_xlink_show(attr.unescape_value()?.as_bytes())?);
            }
            attr => {
                dump_unused(
                    "read_metadata_hyperlink_behaviour",
                    tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    Ok(hyperlink_behaviour)
}

fn read_metadata_document_statistics(
    tag: &BytesStart<'_>,
) -> Result<MetaDocumentStatistics, OdsError> {
    let mut document_statistics = MetaDocumentStatistics::default();

    for attr in tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"meta:cell-count" => {
                document_statistics.cell_count = parse_u32(attr.unescape_value()?.as_bytes())?;
            }
            attr if attr.key.as_ref() == b"meta:object-count" => {
                document_statistics.object_count = parse_u32(attr.unescape_value()?.as_bytes())?;
            }
            attr if attr.key.as_ref() == b"meta:ole-object-count" => {
                document_statistics.ole_object_count =
                    parse_u32(attr.unescape_value()?.as_bytes())?;
            }
            attr if attr.key.as_ref() == b"meta:table-count" => {
                document_statistics.table_count = parse_u32(attr.unescape_value()?.as_bytes())?;
            }
            attr => {
                dump_unused(
                    "read_metadata_document_statistics",
                    tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    Ok(document_statistics)
}

fn read_metadata_user_defined(
    bs: &mut BufStack,
    tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<MetaUserDefined, OdsError> {
    let mut user_defined = MetaUserDefined::default();
    let mut value_type = None;
    for attr in tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"meta:name" => {
                user_defined.name = attr.unescape_value()?.to_string();
            }
            attr if attr.key.as_ref() == b"meta:value-type" => {
                value_type = Some(match attr.unescape_value()?.as_ref() {
                    "boolean" => "boolean",
                    "date" => "date",
                    "float" => "float",
                    "time" => "time",
                    _ => "string",
                });
            }
            attr => {
                dump_unused("read_meta_user_defined", tag.name().as_ref(), &attr)?;
            }
        }
    }

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!("read_meta_user_defined {:?}", evt);
        }

        match &evt {
            Event::End(xml_tag) if xml_tag.name() == tag.name() => {
                break;
            }
            Event::Text(v) => {
                user_defined.value = match value_type {
                    Some("boolean") => MetaValue::Boolean(parse_bool(v.unescape()?.as_bytes())?),
                    Some("date") => MetaValue::Datetime(parse_datetime(v.unescape()?.as_bytes())?),
                    Some("float") => MetaValue::Float(parse_f64(v.unescape()?.as_bytes())?),
                    Some("time") => {
                        MetaValue::TimeDuration(parse_duration(v.unescape()?.as_bytes())?)
                    }
                    _ => MetaValue::String(v.unescape()?.to_string()),
                };
            }
            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_meta_user_defined", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(user_defined)
}

// Parse a metadata value.
fn read_metadata_value<T>(
    bs: &mut BufStack,
    tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
    parse: fn(&[u8]) -> Result<T, OdsError>,
    default: fn() -> T,
) -> Result<T, OdsError> {
    let mut buf = bs.pop();
    let mut value = None;
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!("read_metadata_value {:?}", evt);
        }

        match &evt {
            Event::End(xml_tag) if xml_tag.name() == tag.name() => {
                break;
            }
            Event::Text(v) => {
                value = Some(parse(v.unescape()?.as_bytes())?);
            }
            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_metadata_value", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(value.unwrap_or(default()))
}

fn read_ods_settings(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_settings {:?}", evt);
        }

        match &evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:document-settings" => {
                let (_, xmlns) = read_namespaces_and_version(xml_tag)?;
                book.xmlns.insert("settings.xml".to_string(), xmlns);
            }
            Event::End(xml_tag) if xml_tag.name().as_ref() == b"office:document-settings" => {}

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"office:settings" => {
                read_office_settings(bs, book, xml)?;
            }

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_settings", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// read the automatic-styles tag
fn read_office_settings(
    bs: &mut BufStack,
    book: &mut WorkBook,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(), OdsError> {
    let mut config = Config::new();

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_office_settings {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-set" => {
                let (name, set) = read_config_item_set(bs, xml_tag, xml)?;
                config.insert(name, set);
            }

            Event::End(e) if e.name().as_ref() == b"office:settings" => {
                break;
            }
            Event::Eof => break,
            _ => {
                dump_unused2("read_office_settings", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    book.config = Detach::new(config);

    Ok(())
}

// read the automatic-styles tag
fn read_config_item_set(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(String, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_set = ConfigItem::new_set();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"config:name" => {
                name = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused("read_config_item_set", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    let name = if let Some(name) = name {
        name
    } else {
        return Err(OdsError::Ods("config-item-set without name".to_string()));
    };

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_office_item_set {:?}", evt);
        }
        match &evt {
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"config:config-item" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item" => {
                let (name, val) = read_config_item(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-set" => {
                let (name, val) = read_config_item_set(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(xml_tag)
                if xml_tag.name().as_ref() == b"config:config-item-map-indexed" =>
            {
                let (name, val) = read_config_item_map_indexed(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-map-named" => {
                let (name, val) = read_config_item_map_named(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::End(e) if e.name().as_ref() == b"config:config-item-set" => {
                break;
            }
            Event::Eof => break,
            _ => {
                dump_unused2("read_config_item_set", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok((name, config_set))
}

// read the automatic-styles tag
fn read_config_item_map_indexed(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(String, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_vec = ConfigItem::new_vec();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"config:name" => {
                name = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused(
                    "read_config_item_map_indexed",
                    super_tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    let name = if let Some(name) = name {
        name
    } else {
        return Err(OdsError::Ods(
            "config-item-map-indexed without name".to_string(),
        ));
    };

    let mut index = 0;

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_office_item_set {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-map-entry" => {
                let (_, entry) = read_config_item_map_entry(bs, xml_tag, xml)?;
                config_vec.insert(index.to_string(), entry);
                index += 1;
            }
            Event::End(e) if e.name().as_ref() == b"config:config-item-map-indexed" => {
                break;
            }
            Event::Eof => break,
            _ => {
                dump_unused2("read_config_item_map_indexed", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok((name, config_vec))
}

// read the automatic-styles tag
fn read_config_item_map_named(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(String, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_map = ConfigItem::new_map();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"config:name" => {
                name = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused(
                    "read_config_item_map_named",
                    super_tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    let name = if let Some(name) = name {
        name
    } else {
        return Err(OdsError::Ods(
            "config-item-map-named without name".to_string(),
        ));
    };

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_config_item_map_named {:?}", evt);
        }
        match &evt {
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-map-entry" => {
                let (name, entry) = read_config_item_map_entry(bs, xml_tag, xml)?;

                let name = if let Some(name) = name {
                    name
                } else {
                    return Err(OdsError::Ods(
                        "config-item-map-entry without name".to_string(),
                    ));
                };

                config_map.insert(name, entry);
            }
            Event::End(e) if e.name().as_ref() == b"config:config-item-map-named" => {
                break;
            }
            Event::Eof => break,
            _ => {
                dump_unused2("read_config_item_map_named", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok((name, config_map))
}

// read the automatic-styles tag
fn read_config_item_map_entry(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(Option<String>, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_set = ConfigItem::new_entry();

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"config:name" => {
                name = Some(attr.unescape_value()?.to_string());
            }
            attr => {
                dump_unused(
                    "read_config_item_map_entry",
                    super_tag.name().as_ref(),
                    &attr,
                )?;
            }
        }
    }

    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        if DUMP_XML {
            println!(" read_config_item_map_entry {:?}", evt);
        }
        match &evt {
            Event::Empty(xml_tag) if xml_tag.name().as_ref() == b"config:config-item" => {}
            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item" => {
                let (name, val) = read_config_item(bs, xml_tag, xml)?;
                config_set.insert(name, ConfigItem::from(val));
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-set" => {
                let (name, val) = read_config_item_set(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }

            Event::Start(xml_tag)
                if xml_tag.name().as_ref() == b"config:config-item-map-indexed" =>
            {
                let (name, val) = read_config_item_map_indexed(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }

            Event::Start(xml_tag) if xml_tag.name().as_ref() == b"config:config-item-map-named" => {
                let (name, val) = read_config_item_map_named(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::End(e) if e.name().as_ref() == b"config:config-item-map-entry" => {
                break;
            }

            Event::Eof => break,
            _ => {
                dump_unused2("read_config_item_map_entry", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok((name, config_set))
}

// read the automatic-styles tag
fn read_config_item(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<(String, ConfigValue), OdsError> {
    #[derive(PartialEq)]
    enum ConfigValueType {
        None,
        Base64Binary,
        Boolean,
        DateTime,
        Double,
        Int,
        Long,
        Short,
        String,
    }

    let mut name = None;
    let mut val_type = ConfigValueType::None;
    let mut config_val = None;

    for attr in super_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key.as_ref() == b"config:name" => {
                name = Some(attr.unescape_value()?.to_string());
            }
            attr if attr.key.as_ref() == b"config:type" => {
                val_type = match attr.value.as_ref() {
                    b"base64Binary" => ConfigValueType::Base64Binary,
                    b"boolean" => ConfigValueType::Boolean,
                    b"datetime" => ConfigValueType::DateTime,
                    b"double" => ConfigValueType::Double,
                    b"int" => ConfigValueType::Int,
                    b"long" => ConfigValueType::Long,
                    b"short" => ConfigValueType::Short,
                    b"string" => ConfigValueType::String,
                    x => {
                        return Err(OdsError::Ods(format!(
                            "unknown config:type {}",
                            from_utf8(x)?
                        )));
                    }
                };
            }
            attr => {
                dump_unused("read_config_item", super_tag.name().as_ref(), &attr)?;
            }
        }
    }

    let name = if let Some(name) = name {
        name
    } else {
        return Err(OdsError::Ods(
            "config value without config:name".to_string(),
        ));
    };

    if val_type == ConfigValueType::None {
        return Err(OdsError::Ods(
            "config value without config:type".to_string(),
        ));
    };

    let mut value = bs.pop();
    let mut buf = bs.pop();
    loop {
        let evt = xml.read_event_into(&mut buf)?;
        match &evt {
            Event::Text(txt) => {
                value.write_all(txt.unescape()?.as_bytes())?;
            }
            Event::End(e) if e.name().as_ref() == b"config:config-item" => {
                let value = <Cow<'_, [u8]> as From<&Vec<u8>>>::from(value.as_ref());
                match val_type {
                    ConfigValueType::None => {}
                    ConfigValueType::Base64Binary => {
                        config_val =
                            Some(ConfigValue::Base64Binary(from_utf8(&value)?.to_string()));
                    }
                    ConfigValueType::Boolean => {
                        config_val = Some(ConfigValue::Boolean(parse_bool(&value)?));
                    }
                    ConfigValueType::DateTime => {
                        config_val = Some(ConfigValue::DateTime(parse_datetime(&value)?));
                    }
                    ConfigValueType::Double => {
                        config_val = Some(ConfigValue::Double(parse_f64(&value)?));
                    }
                    ConfigValueType::Int => {
                        config_val = Some(ConfigValue::Int(parse_i32(&value)?));
                    }
                    ConfigValueType::Long => {
                        config_val = Some(ConfigValue::Long(parse_i64(&value)?));
                    }
                    ConfigValueType::Short => {
                        config_val = Some(ConfigValue::Short(parse_i16(&value)?));
                    }
                    ConfigValueType::String => {
                        config_val =
                            Some(ConfigValue::String(from_utf8(value.as_ref())?.to_string()));
                    }
                }
                break;
            }
            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_config_item", &evt)?;
            }
        }

        if DUMP_XML {
            println!(" read_config_item {:?}", evt);
        }
        buf.clear();
    }
    bs.push(buf);
    bs.push(value);

    let config_val = if let Some(config_val) = config_val {
        config_val
    } else {
        return Err(OdsError::Ods("config-item without value???".to_string()));
    };

    Ok((name, config_val))
}

// Reads a part of the XML as XmlTag's.
fn read_xml(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<XmlTag, OdsError> {
    let mut stack = Vec::new();

    let mut tag = XmlTag::new(xml.decoder().decode(super_tag.name().as_ref())?);
    copy_attr2(tag.attrmap_mut(), super_tag)?;
    stack.push(tag);

    if !empty_tag {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_xml {:?}", evt);
            }
            match &evt {
                Event::Start(xmlbytes) => {
                    let mut tag = XmlTag::new(xml.decoder().decode(xmlbytes.name().as_ref())?);
                    copy_attr2(tag.attrmap_mut(), xmlbytes)?;
                    stack.push(tag);
                }

                Event::End(xmlbytes) => {
                    if xmlbytes.name() == super_tag.name() {
                        break;
                    } else {
                        let tag = stack.pop().unwrap();
                        if let Some(parent) = stack.last_mut() {
                            parent.add_tag(tag);
                        } else {
                            unreachable!()
                        }
                    }
                }

                Event::Empty(xmlbytes) => {
                    let mut emptytag = XmlTag::new(xml.decoder().decode(xmlbytes.name().as_ref())?);
                    copy_attr2(emptytag.attrmap_mut(), xmlbytes)?;

                    if let Some(parent) = stack.last_mut() {
                        parent.add_tag(emptytag);
                    } else {
                        unreachable!()
                    }
                }

                Event::Text(xmlbytes) => {
                    if let Some(parent) = stack.last_mut() {
                        parent.add_text(
                            xml.decoder()
                                .decode(xmlbytes.unescape().unwrap().as_bytes())
                                .unwrap(),
                        );
                    } else {
                        unreachable!()
                    }
                }

                Event::Eof => {
                    break;
                }

                _ => {
                    dump_unused2("read_xml", &evt)?;
                }
            }
            buf.clear();
        }

        bs.push(buf);
    }

    assert_eq!(stack.len(), 1);
    Ok(stack.pop().unwrap())
}

fn read_text_or_tag(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    xml: &mut OdsXmlReader<'_>,
) -> Result<TextContent2, OdsError> {
    let mut stack = Vec::new();
    let mut cellcontent = TextContent2::Empty;

    // The toplevel element is passed in with the xml_tag.
    // It is only created if there are further xml tags in the
    // element. If there is only text this is not needed.
    let create_toplevel = |t: Option<String>| -> Result<XmlTag, OdsError> {
        // No parent tag on the stack. Create the parent.
        let mut toplevel = XmlTag::new(from_utf8(super_tag.name().as_ref())?);
        copy_attr2(toplevel.attrmap_mut(), super_tag)?;
        if let Some(t) = t {
            toplevel.add_text(t);
        }
        Ok(toplevel)
    };

    if !empty_tag {
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_xml {:?}", evt);
            }
            match &evt {
                Event::Start(xmlbytes) => {
                    match cellcontent {
                        TextContent2::Empty => {
                            stack.push(create_toplevel(None)?);
                        }
                        TextContent2::Text(old_txt) => {
                            stack.push(create_toplevel(Some(old_txt))?);
                        }
                        TextContent2::Xml(parent) => {
                            stack.push(parent);
                        }
                        TextContent2::XmlVec(_) => {
                            unreachable!()
                        }
                    }

                    // Set the new tag.
                    let mut new_tag = XmlTag::new(xml.decoder().decode(xmlbytes.name().as_ref())?);
                    copy_attr2(new_tag.attrmap_mut(), xmlbytes)?;
                    cellcontent = TextContent2::Xml(new_tag)
                }
                Event::Empty(xmlbytes) => {
                    match cellcontent {
                        TextContent2::Empty => {
                            stack.push(create_toplevel(None)?);
                        }
                        TextContent2::Text(txt) => {
                            stack.push(create_toplevel(Some(txt))?);
                        }
                        TextContent2::Xml(parent) => {
                            stack.push(parent);
                        }
                        TextContent2::XmlVec(_) => {
                            unreachable!()
                        }
                    }
                    if let Some(mut parent) = stack.pop() {
                        // Create the tag and append it immediately to the parent.
                        let mut emptytag =
                            XmlTag::new(xml.decoder().decode(xmlbytes.name().as_ref())?);
                        copy_attr2(emptytag.attrmap_mut(), xmlbytes)?;
                        parent.add_tag(emptytag);

                        cellcontent = TextContent2::Xml(parent);
                    } else {
                        unreachable!()
                    }
                }
                Event::Text(xmlbytes) => {
                    let v = xmlbytes.unescape()?;

                    cellcontent = match cellcontent {
                        TextContent2::Empty => {
                            // Fresh plain text string.
                            TextContent2::Text(v.to_string())
                        }
                        TextContent2::Text(mut old_txt) => {
                            // We have a previous plain text string. Append to it.
                            old_txt.push_str(&v);
                            TextContent2::Text(old_txt)
                        }
                        TextContent2::Xml(mut xml) => {
                            // There is already a tag. Append the text to its children.
                            xml.add_text(v);
                            TextContent2::Xml(xml)
                        }
                        TextContent2::XmlVec(_) => {
                            unreachable!()
                        }
                    };
                }
                Event::End(xmlbytes) if xmlbytes.name() == super_tag.name() => {
                    if !stack.is_empty() {
                        return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(format!(
                            "XML corrupted. Endtag {} occured before all elements are closed: {:?}",
                            from_utf8(super_tag.name().as_ref())?,
                            stack
                        ))));
                    }
                    break;
                }
                Event::End(xmlbytes) => {
                    cellcontent = match cellcontent {
                        TextContent2::Empty | TextContent2::Text(_) => {
                            return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(format!(
                                "XML corrupted. Endtag {} occured without start tag",
                                from_utf8(xmlbytes.name().as_ref())?
                            ))));
                        }
                        TextContent2::Xml(tag) => {
                            if let Some(mut parent) = stack.pop() {
                                parent.add_tag(tag);
                                TextContent2::Xml(parent)
                            } else {
                                return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(
                                    format!(
                                        "XML corrupted. Endtag {} occured without start tag",
                                        from_utf8(xmlbytes.name().as_ref())?
                                    ),
                                )));
                            }
                        }
                        TextContent2::XmlVec(_) => {
                            unreachable!()
                        }
                    }
                }

                Event::Eof => {
                    break;
                }

                _ => {
                    dump_unused2("read_text_or_tag", &evt)?;
                }
            }
        }
        bs.push(buf);
    }

    Ok(cellcontent)
}

/// Read simple text content.
/// Fail on any tag other than the end-tag to the supertag.
fn read_text<T, E>(
    bs: &mut BufStack,
    super_tag: &BytesStart<'_>,
    empty_tag: bool,
    parse: fn(&[u8]) -> Result<T, E>,
    xml: &mut OdsXmlReader<'_>,
) -> Result<Option<T>, OdsError>
where
    OdsError: From<E>,
{
    if empty_tag {
        Ok(None)
    } else {
        let mut r = Vec::new();
        let mut buf = bs.pop();
        loop {
            let evt = xml.read_event_into(&mut buf)?;
            if DUMP_XML {
                println!(" read_text {:?}", evt);
            }
            match &evt {
                Event::Text(xml_tag) => {
                    r.extend_from_slice(xml_tag.as_ref());
                }
                Event::End(xml_tag) if xml_tag.name() == super_tag.name() => {
                    break;
                }
                Event::Empty(xml_tag) | Event::Start(xml_tag) => {
                    return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(
                        from_utf8(xml_tag.as_ref())?.to_string(),
                    )))
                }
                Event::End(xml_tag) => {
                    return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(
                        from_utf8(xml_tag.as_ref())?.to_string(),
                    )))
                }

                Event::Eof => {
                    break;
                }

                _ => {
                    dump_unused2("read_text", &evt)?;
                }
            }
        }
        bs.push(buf);

        Ok(Some(parse(&r)?))
    }
}

#[inline(always)]
fn dump_unused(func: &str, tag: &[u8], attr: &Attribute<'_>) -> Result<(), OdsError> {
    if DUMP_UNUSED {
        let tag = from_utf8(tag)?;
        let key = from_utf8(attr.key.as_ref())?;
        let value = from_utf8(attr.value.as_ref())?;
        println!("unused attr: {} '{}' ({}:{})", func, tag, key, value);
    }
    Ok(())
}

#[inline(always)]
fn dump_unused2(func: &str, evt: &Event<'_>) -> Result<(), OdsError> {
    if DUMP_UNUSED {
        match &evt {
            Event::Text(text) => {
                if !text.unescape()?.trim().is_empty() {
                    println!("unused text: {} ({:?})", func, evt);
                }
            }
            _ => {
                println!("unused event: {} ({:?})", func, evt);
            }
        }
    }
    Ok(())
}
