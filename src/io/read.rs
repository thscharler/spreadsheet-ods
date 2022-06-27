use std::convert::{TryFrom, TryInto};
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;

use chrono::{Duration, NaiveDateTime};
use quick_xml::events::{BytesStart, Event};
use zip::read::ZipFile;
use zip::ZipArchive;

use crate::attrmap2::AttrMap2;
use crate::condition::{Condition, ValueCondition};
use crate::config::{Config, ConfigItem, ConfigItemType, ConfigValue};
use crate::ds::bufstack::BufStack;
use crate::ds::detach::Detach;
use crate::error::OdsError;
use crate::format::{FormatPart, FormatPartType};
use crate::io::parse::{
    parse_bool, parse_currency, parse_datetime, parse_duration, parse_f64, parse_i16, parse_i32,
    parse_i64, parse_string, parse_u32, parse_visibility,
};
use crate::io::{DUMP_UNUSED, DUMP_XML};
use crate::refs::{parse_cellranges, parse_cellref};
use crate::style::stylemap::StyleMap;
use crate::style::tabstop::TabStop;
use crate::style::{
    ColStyle, FontFaceDecl, GraphicStyle, HeaderFooter, MasterPage, PageStyle, ParagraphStyle,
    RowStyle, StyleOrigin, StyleUse, TableStyle, TextStyle,
};
use crate::text::{TextP, TextTag};
use crate::validation::{MessageType, Validation, ValidationError, ValidationHelp};
use crate::xmltree::{XmlContent, XmlTag};
use crate::{
    CellData, CellStyle, ColRange, Length, RowRange, Sheet, SplitMode, Value, ValueFormat,
    ValueType, Visibility, WorkBook,
};
use quick_xml::events::attributes::Attribute;
use std::borrow::Cow;
use std::str::from_utf8;

/// Reads an ODS-file from a buffer
pub fn read_ods_buf(buf: &[u8]) -> Result<WorkBook, OdsError> {
    let zip = ZipArchive::new(Cursor::new(buf))?;
    read_ods_impl(zip)
}

/// Reads an ODS-file.
pub fn read_ods<P: AsRef<Path>>(path: P) -> Result<WorkBook, OdsError> {
    let file = File::open(path.as_ref())?;
    let zip = ZipArchive::new(file)?;
    read_ods_impl(zip)
}

/// Reads an ODS-file.
fn read_ods_impl<R: Read + Seek>(mut zip: ZipArchive<R>) -> Result<WorkBook, OdsError> {
    let mut book = WorkBook::new_empty();
    let mut bufstack = BufStack::new();

    read_content(&mut bufstack, &mut book, &mut zip.by_name("content.xml")?)?;
    read_styles(&mut bufstack, &mut book, &mut zip.by_name("styles.xml")?)?;
    // may not exist.
    if let Ok(mut z) = zip.by_name("settings.xml") {
        read_settings(&mut bufstack, &mut book, &mut z)?;
    } else {
        book.config = default_settings();
    }

    // read all extras.
    read_filebuf(&mut book, &mut zip)?;

    // We do some data duplication here, to make everything easier to use.
    calc_derived(&mut book)?;

    Ok(book)
}

// Loads all unprocessed files as byte blobs into a buffer.
fn read_filebuf<R: Read + Seek>(
    book: &mut WorkBook,
    zip: &mut ZipArchive<R>,
) -> Result<(), OdsError> {
    for idx in 0..zip.len() {
        let mut ze = zip.by_index(idx)?;

        // These three are always interpreted and rewritten from scratch.
        // They have their own mechanism to cope with unknown data.
        if !matches!(ze.name(), "settings.xml" | "styles.xml" | "content.xml") {
            if ze.is_dir() {
                book.filebuf.push_dir(ze.name());
            } else if ze.is_file() {
                let mut buf = Vec::new();
                ze.read_to_end(&mut buf)?;
                book.filebuf.push_file(ze.name(), buf);
            }
        }
    }

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
fn read_content(
    bs: &mut BufStack,
    book: &mut WorkBook,
    zip_file: &mut ZipFile<'_>,
) -> Result<(), OdsError> {
    // xml parser
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    // Do not trim text data. All text read contains significant whitespace.
    // The rest is ignored anyway.
    //
    // xml.trim_text(true);

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_content {:?}", evt);
        }
        match evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:body"
                || xml_tag.name() == b"office:spreadsheet" => {
                // noop
            }
            Event::End(xml_tag)
            if xml_tag.name() == b"office:body"
                || xml_tag.name() == b"office:spreadsheet" => {
                // noop
            }
            Event::Start(xml_tag)
            if xml_tag.name() == b"office:document-content" => {
                for attr in xml_tag.attributes().with_checks(false) {
                    match attr? {
                        attr if attr.key == b"office:version" => {
                            book.set_version(parse_string(&attr.value)?);
                        }
                        _ => {
                            // noop
                        }
                    }
                }
            }
            Event::End(xml_tag)
            if xml_tag.name() == b"office:document-content" => {
                // noop
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:font-face-decls" =>
                read_fonts(bs, book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:styles" =>
                read_styles_tag(bs, book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:automatic-styles" =>
                read_auto_styles(bs, book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:master-styles" =>
                read_master_styles(bs, book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:content-validations" =>
                read_validations(bs, book, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table" =>
                book.push_sheet(read_table(bs, &mut xml, xml_tag)?),

            Event::Empty(xml_tag) |
            Event::Start(xml_tag)
            if /* prelude */ xml_tag.name() == b"office:scripts" ||
                xml_tag.name() == b"table:tracked-changes" ||
                xml_tag.name() == b"text:variable-decls" ||
                xml_tag.name() == b"text:sequence-decls" ||
                xml_tag.name() == b"text:user-field-decls" ||
                xml_tag.name() == b"text:dde-connection-decls" ||
                // xml_tag.name() == b"text:alphabetical-index-auto-mark-file" ||
                xml_tag.name() == b"table:calculation-settings" ||
                xml_tag.name() == b"table:label-ranges" ||
                /* epilogue */
                xml_tag.name() == b"table:named-expressions" ||
                xml_tag.name() == b"table:database-ranges" ||
                xml_tag.name() == b"table:data-pilot-tables" ||
                xml_tag.name() == b"table:consolidation" ||
                xml_tag.name() == b"table:dde-links" => {
                let v = read_xml(bs, xml_tag.name(), &mut xml, &xml_tag, empty_tag)?;
                book.extra.push(v);
            }

            Event::End(xml_tag)
            if /* prelude */ xml_tag.name() == b"office:scripts" ||
                xml_tag.name() == b"table:tracked-changes" ||
                xml_tag.name() == b"text:variable-decls" ||
                xml_tag.name() == b"text:sequence-decls" ||
                xml_tag.name() == b"text:user-field-decls" ||
                xml_tag.name() == b"text:dde-connection-decls" ||
                // xml_tag.name() == b"text:alphabetical-index-auto-mark-file" ||
                xml_tag.name() == b"table:calculation-settings" ||
                xml_tag.name() == b"table:label-ranges" ||
                /* epilogue */
                xml_tag.name() == b"table:named-expressions" ||
                xml_tag.name() == b"table:database-ranges" ||
                xml_tag.name() == b"table:data-pilot-tables" ||
                xml_tag.name() == b"table:consolidation" ||
                xml_tag.name() == b"table:dde-links" => {
                // noop
            }

            Event::Eof => {
                break;
            }
            _ => {
                dump_unused2("read_content", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

// Reads the table.
fn read_table(
    bs: &mut BufStack,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: BytesStart<'_>,
) -> Result<Sheet, OdsError> {
    let mut sheet = Sheet::new("");

    read_table_attr(&mut sheet, xml_tag)?;

    // Position within table-columns
    let mut table_col: u32 = 0;

    // Cell position
    let mut row: u32 = 0;
    let mut col: u32 = 0;

    // Rows can be repeated. In reality only empty ones ever are.
    let mut row_repeat: u32 = 1;
    let mut rowstyle: Option<String> = None;
    let mut row_cellstyle: Option<String> = None;
    let mut row_visible: Visibility = Default::default();

    let mut col_range_from = 0;
    let mut row_range_from = 0;

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_table {:?}", evt);
        }
        match evt {
            Event::End(xml_tag)
            if xml_tag.name() == b"table:table" => {
                break;
            }

            Event::Start(xml_tag) |
            Event::Empty(xml_tag)
            if /* prelude */ xml_tag.name() == b"table:title" ||
                xml_tag.name() == b"table:desc" ||
                xml_tag.name() == b"table:table-source" ||
                xml_tag.name() == b"office:dde-source" ||
                xml_tag.name() == b"table:scenario" ||
                xml_tag.name() == b"office:forms" ||
                xml_tag.name() == b"table:shapes" ||
                /* epilogue */
                xml_tag.name() == b"table:named-expressions" ||
                xml_tag.name() == b"calcext:conditional-formats" => {
                sheet.extra.push(read_xml(bs, xml_tag.name(), xml, &xml_tag, empty_tag)?);
            }

            Event::End(xml_tag)
            if /* prelude */ xml_tag.name() == b"table:title" ||
                xml_tag.name() == b"table:desc" ||
                xml_tag.name() == b"table:table-source" ||
                xml_tag.name() == b"office:dde-source" ||
                xml_tag.name() == b"table:scenario" ||
                xml_tag.name() == b"office:forms" ||
                xml_tag.name() == b"table:shapes" ||
                /* epilogue */
                xml_tag.name() == b"table:named-expressions" ||
                xml_tag.name() == b"calcext:conditional-formats" => {}

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-header-columns" => {
                col_range_from = table_col;
            }

            Event::End(xml_tag)
            if xml_tag.name() == b"table:table-header-columns" => {
                sheet.header_cols = Some(ColRange::new(col_range_from, table_col - 1));
            }

            Event::Empty(xml_tag)
            if xml_tag.name() == b"table:table-column" => {
                table_col = read_table_col_attr(&mut sheet, table_col,  &xml_tag)?;
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-header-rows" => {
                row_range_from = row;
            }

            Event::End(xml_tag)
            if xml_tag.name() == b"table:table-header-rows" => {
                sheet.header_rows = Some(RowRange::new(row_range_from, row - 1));
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-row" => {
                let (repeat, style, cellstyle, visible) = read_table_row_attr( xml_tag)?;
                row_repeat = repeat;
                rowstyle = style;
                row_cellstyle = cellstyle;
                row_visible = visible;
            }

            Event::End(xml_tag)
            if xml_tag.name() == b"table:table-row" => {
                if row_repeat > 1 {
                    sheet.set_row_repeat(row, row_repeat);
                }
                if let Some(rowstyle) = rowstyle {
                    sheet.set_rowstyle(row, &rowstyle.into());
                }
                rowstyle = None;
                if let Some(row_cellstyle) = row_cellstyle {
                    sheet.set_row_cellstyle(row, &row_cellstyle.into());
                }
                row_cellstyle = None;
                if row_visible != Visibility::Visible {
                    sheet.set_row_visible(row, row_visible);
                }
                row_visible = Default::default();

                row += row_repeat;
                col = 0;
                row_repeat = 1;
            }

            Event::Empty(xml_tag)
            if xml_tag.name() == b"table:table-cell" || xml_tag.name() == b"table:covered-table-cell" => {
                col = read_empty_table_cell(&mut sheet, row, col, xml_tag)?;
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-cell" || xml_tag.name() == b"table:covered-table-cell" => {
                col = read_table_cell2(bs, &mut sheet, row, col, xml, xml_tag)?;
            }

            _ => {
                dump_unused2("read_table", &evt)?;
            }
        }
        buf.clear();
    }
    bs.push(buf);

    Ok(sheet)
}

// Reads the table attributes.
fn read_table_attr(sheet: &mut Sheet, xml_tag: BytesStart<'_>) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:name" => {
                sheet.set_name(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:style-name" => {
                sheet.set_style(&parse_string(&attr.value)?.into());
            }
            attr if attr.key == b"table:print" => {
                sheet.set_print(parse_bool(&attr.value)?);
            }
            attr if attr.key == b"table:display" => {
                sheet.set_display(parse_bool(&attr.value)?);
            }
            attr if attr.key == b"table:print-ranges" => {
                let v = attr.unescaped_value()?;
                let mut pos = 0usize;
                sheet.print_ranges = parse_cellranges(from_utf8(v.as_ref())?, &mut pos)?;
            }
            attr => {
                dump_unused("read_table_attr", xml_tag.name(), &attr)?;
            }
        }
    }

    Ok(())
}

// Reads table-row attributes. Returns the repeat-count.
fn read_table_row_attr(
    xml_tag: BytesStart<'_>,
) -> Result<(u32, Option<String>, Option<String>, Visibility), OdsError> {
    let mut row_repeat: u32 = 1;
    let mut row_visible: Visibility = Default::default();
    let mut rowstyle: Option<String> = None;
    let mut row_cellstyle: Option<String> = None;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            // table:default-cell-style-name 19.615, table:visibility 19.749 and xml:id 19.914.
            attr if attr.key == b"table:number-rows-repeated" => {
                row_repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:style-name" => {
                rowstyle = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:default-cell-style-name" => {
                row_cellstyle = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:visibility" => {
                row_visible = parse_visibility(&attr.value)?;
            }
            attr => {
                dump_unused("read_table_row_attr", xml_tag.name(), &attr)?;
            }
        }
    }

    Ok((row_repeat, rowstyle, row_cellstyle, row_visible))
}

// Reads the table-column attributes. Creates as many copies as indicated.
fn read_table_col_attr(
    sheet: &mut Sheet,
    mut table_col: u32,
    xml_tag: &BytesStart<'_>,
) -> Result<u32, OdsError> {
    let mut style = None;
    let mut cellstyle = None;
    let mut repeat: u32 = 1;
    let mut visible: Visibility = Default::default();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-columns-repeated" => {
                repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:style-name" => {
                style = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:default-cell-style-name" => {
                cellstyle = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:visibility" => {
                visible = parse_visibility(&attr.value)?;
            }
            attr => {
                dump_unused("read_table_col_attr", xml_tag.name(), &attr)?;
            }
        }
    }

    while repeat > 0 {
        if let Some(style) = &style {
            sheet.set_colstyle(table_col, &style.into());
        }
        if let Some(cellstyle) = &cellstyle {
            sheet.set_col_cellstyle(table_col, &cellstyle.into());
        }
        sheet.set_col_visible(table_col, visible);
        table_col += 1;
        repeat -= 1;
    }

    Ok(table_col)
}

#[derive(Debug)]
#[allow(variant_size_differences)]
enum TextContent {
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
    val_currency: Option<[u8; 3]>,

    content: TextContent,
}

fn read_table_cell2(
    bs: &mut BufStack,
    sheet: &mut Sheet,
    row: u32,
    mut col: u32,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: BytesStart<'_>,
) -> Result<u32, OdsError> {
    // Current cell tag
    let tag_name = xml_tag.name();

    let mut cell_repeat: u32 = 1;

    let mut cell = CellData {
        value: Default::default(),
        formula: None,
        style: None,
        validation_name: None,
        span: Default::default(),
    };

    let mut tc = ReadTableCell2 {
        val_type: ValueType::Empty,
        val_datetime: None,
        val_duration: None,
        val_float: None,
        val_bool: None,
        val_string: None,
        val_currency: None,
        content: TextContent::Empty,
    };

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-columns-repeated" => {
                cell_repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:number-rows-spanned" => {
                cell.span.row_span = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:number-columns-spanned" => {
                cell.span.col_span = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:content-validation-name" => {
                cell.validation_name = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"calcext:value-type" => {
                // not used. office:value-type seems to be good enough.
            }
            attr if attr.key == b"office:value-type" => {
                tc.val_type = match attr.value.as_ref() {
                    b"string" => ValueType::Text,
                    b"float" => ValueType::Number,
                    b"percentage" => ValueType::Percentage,
                    b"date" => ValueType::DateTime,
                    b"time" => ValueType::TimeDuration,
                    b"boolean" => ValueType::Boolean,
                    b"currency" => ValueType::Currency,
                    other => return Err(OdsError::Parse(format!("Unknown cell-type {:?}", other))),
                }
            }
            attr if attr.key == b"office:date-value" => {
                tc.val_datetime = Some(parse_datetime(&attr.value)?);
            }
            attr if attr.key == b"office:time-value" => {
                tc.val_duration = Some(parse_duration(&attr.value)?);
            }
            attr if attr.key == b"office:value" => {
                tc.val_float = Some(parse_f64(&attr.value)?);
            }
            attr if attr.key == b"office:boolean-value" => {
                tc.val_bool = Some(parse_bool(&attr.value)?);
            }
            attr if attr.key == b"office:string-value" => {
                tc.val_string = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"office:currency" => {
                tc.val_currency = Some(parse_currency(&attr.value)?);
            }
            attr if attr.key == b"table:formula" => {
                cell.formula = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:style-name" => {
                cell.style = Some(parse_string(attr.value.as_ref())?);
            }
            attr => {
                dump_unused("read_table_cell2", xml_tag.name(), &attr)?;
            }
        }
    }

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_table_cell {:?}", evt);
        }
        match evt {
            Event::Start(xml_tag) if xml_tag.name() == b"text:p" => {
                let new_txt = read_text_or_tag(bs, b"text:p", xml, &xml_tag, false)?;
                tc = append_text(new_txt, tc);
            }
            Event::Empty(xml_tag) if xml_tag.name() == b"text:p" => {
                // noop
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

fn append_text(new_txt: TextContent, mut tc: ReadTableCell2) -> ReadTableCell2 {
    // There can be multiple text:p elements within the cell.
    tc.content = match tc.content {
        TextContent::Empty => new_txt,
        TextContent::Text(txt) => {
            // Have a destructured text:p from before.
            // Wrap up and create list.
            let p = TextP::new().text(txt).into_xmltag();
            let mut vec = vec![p];

            match new_txt {
                TextContent::Empty => {}
                TextContent::Text(txt) => {
                    let p2 = TextP::new().text(txt).into_xmltag();
                    vec.push(p2);
                }
                TextContent::Xml(xml) => {
                    vec.push(xml);
                }
                TextContent::XmlVec(_) => {
                    unreachable!();
                }
            }
            TextContent::XmlVec(vec)
        }
        TextContent::Xml(xml) => {
            let mut vec = vec![xml];
            match new_txt {
                TextContent::Empty => {}
                TextContent::Text(txt) => {
                    let p2 = TextP::new().text(txt).into_xmltag();
                    vec.push(p2);
                }
                TextContent::Xml(xml) => {
                    vec.push(xml);
                }
                TextContent::XmlVec(_) => {
                    unreachable!();
                }
            }
            TextContent::XmlVec(vec)
        }
        TextContent::XmlVec(mut vec) => {
            match new_txt {
                TextContent::Empty => {}
                TextContent::Text(txt) => {
                    let p2 = TextP::new().text(txt).into_xmltag();
                    vec.push(p2);
                }
                TextContent::Xml(xml) => {
                    vec.push(xml);
                }
                TextContent::XmlVec(_) => {
                    unreachable!();
                }
            }
            TextContent::XmlVec(vec)
        }
    };

    tc
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
                return Err(OdsError::Parse("no boolean value".to_string()));
            }
        }
        ValueType::Number => {
            if let Some(v) = tc.val_float {
                cell.value = Value::Number(v);
            } else {
                return Err(OdsError::Parse("no float value".to_string()));
            }
        }
        ValueType::Percentage => {
            if let Some(v) = tc.val_float {
                cell.value = Value::Percentage(v);
            } else {
                return Err(OdsError::Parse("no float value".to_string()));
            }
        }
        ValueType::Currency => {
            if let Some(v) = tc.val_float {
                if let Some(c) = tc.val_currency {
                    cell.value = Value::Currency(v, c);
                } else {
                    return Err(OdsError::Parse("no currency value".to_string()));
                }
            } else {
                return Err(OdsError::Parse("no float value".to_string()));
            }
        }
        ValueType::Text => {
            if let Some(v) = tc.val_string {
                cell.value = Value::Text(v);
            } else {
                match tc.content {
                    TextContent::Empty => {
                        // noop
                    }
                    TextContent::Text(txt) => {
                        cell.value = Value::Text(txt);
                    }
                    TextContent::Xml(xml) => {
                        cell.value = Value::TextXml(vec![xml]);
                    }
                    TextContent::XmlVec(vec) => {
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
                return Err(OdsError::Parse("no datetime value".to_string()));
            }
        }
        ValueType::TimeDuration => {
            if let Some(v) = tc.val_duration {
                cell.value = Value::TimeDuration(v);
            } else {
                return Err(OdsError::Parse("no duration value".to_string()));
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
    xml_tag: BytesStart<'_>,
) -> Result<u32, OdsError> {
    let mut cell = None;
    // Default advance is one column.
    let mut cell_repeat = 1;
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-columns-repeated" => {
                cell_repeat = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:formula" => {
                cell.get_or_insert_with(CellData::new).formula = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:style-name" => {
                cell.get_or_insert_with(CellData::new).style = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"table:number-rows-spanned" => {
                cell.get_or_insert_with(CellData::new).span.row_span = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:number-columns-spanned" => {
                cell.get_or_insert_with(CellData::new).span.col_span = parse_u32(&attr.value)?;
            }
            attr if attr.key == b"table:content-validation-name" => {
                cell.get_or_insert_with(CellData::new).validation_name =
                    Some(parse_string(&attr.value)?);
            }

            attr => {
                dump_unused("read_empty_table_cell", xml_tag.name(), &attr)?;
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

// reads a font-face
fn read_fonts(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(), OdsError> {
    let mut font: FontFaceDecl = FontFaceDecl::new_empty();
    font.set_origin(origin);

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_fonts {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                b"style:font-face" => {
                    let name = proc_style_attr(font.attrmap_mut(), xml_tag)?;
                    font.set_name(name);

                    book.add_font(font);

                    font = FontFaceDecl::new_empty();
                    font.set_origin(StyleOrigin::Content);
                }
                _ => {
                    dump_unused2("read_fonts", &evt)?;
                }
            },
            Event::End(ref e) => {
                if e.name() == b"office:font-face-decls" {
                    break;
                }
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
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
) -> Result<(), OdsError> {
    let mut pl = PageStyle::new_empty();
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                pl.set_name(parse_string(&attr.value)?);
            }
            attr => {
                dump_unused("read_page_style", xml_tag.name(), &attr)?;
            }
        }
    }

    let mut headerstyle = false;
    let mut footerstyle = false;

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_page_layout {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:page-layout-properties" => copy_attr2(pl.style_mut(), xml_tag)?,
                    b"style:header-style" => headerstyle = true,
                    b"style:footer-style" => footerstyle = true,
                    b"style:header-footer-properties" => {
                        if headerstyle {
                            copy_attr2(pl.headerstyle_mut().style_mut(), xml_tag)?;
                        }
                        if footerstyle {
                            copy_attr2(pl.footerstyle_mut().style_mut(), xml_tag)?;
                        }
                    }
                    b"style:background-image" => {
                        // noop for now. sets the background transparent.
                    }
                    _ => {
                        dump_unused2("read_page_layout", &evt)?;
                    }
                }
            }
            Event::Text(_) => (),
            Event::End(ref end) => match end.name() {
                b"style:page-layout" => break,
                b"style:page-layout-properties" => {}
                b"style:header-style" => headerstyle = false,
                b"style:footer-style" => footerstyle = false,
                b"style:header-footer-properties" => {}
                _ => {
                    dump_unused2("read_page_layout", &evt)?;
                }
            },
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
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
) -> Result<(), OdsError> {
    let mut valid = Validation::new();

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_master_styles {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                b"table:content-validation" => {
                    for attr in xml_tag.attributes().with_checks(false) {
                        match attr? {
                            attr if attr.key == b"table:name" => {
                                valid.set_name(parse_string(&attr.value)?);
                            }
                            attr if attr.key == b"table:condition" => {
                                // split off 'of:' prefix
                                let v = parse_string(&attr.value)?;
                                valid.set_condition(Condition::new(v.split_at(3).1));
                            }
                            attr if attr.key == b"table:allow-empty-cell" => {
                                valid.set_allow_empty(parse_bool(&attr.value)?);
                            }
                            attr if attr.key == b"table:base-cell-address" => {
                                // todo: maybe better
                                let v = attr.unescaped_value()?;
                                let mut pos = 0usize;
                                valid.set_base_cell(parse_cellref(
                                    from_utf8(v.as_ref())?,
                                    &mut pos,
                                )?);
                            }
                            attr if attr.key == b"table:display-list" => {
                                valid.set_display(attr.value.as_ref().try_into()?);
                            }
                            attr => {
                                dump_unused("read_validations", xml_tag.name(), &attr)?;
                            }
                        }
                    }

                    if empty_tag {
                        book.add_validation(valid);
                        valid = Validation::new();
                    }
                }
                b"table:error-message" => {
                    let mut ve = ValidationError::new();

                    for attr in xml_tag.attributes().with_checks(false) {
                        match attr? {
                            attr if attr.key == b"table:display" => {
                                ve.set_display(parse_bool(&attr.value)?);
                            }
                            attr if attr.key == b"table:message-type" => {
                                let mt = match attr.value.as_ref() {
                                    b"stop" => MessageType::Error,
                                    b"warning" => MessageType::Warning,
                                    b"information" => MessageType::Info,
                                    _ => {
                                        return Err(OdsError::Parse(format!(
                                            "unknown message-type {}",
                                            parse_string(&attr.value)?
                                        )))
                                    }
                                };
                                ve.set_msg_type(mt);
                            }
                            attr if attr.key == b"table:title" => {
                                ve.set_title(Some(parse_string(&attr.value)?));
                            }
                            attr => {
                                dump_unused("read_validations", xml_tag.name(), &attr)?;
                            }
                        }
                    }
                    let txt =
                        read_text_or_tag(bs, b"table:error-message", xml, xml_tag, empty_tag)?;
                    match txt {
                        TextContent::Empty => {}
                        TextContent::Xml(txt) => {
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
                }
                b"table:help-message" => {
                    let mut vh = ValidationHelp::new();

                    for attr in xml_tag.attributes().with_checks(false) {
                        match attr? {
                            attr if attr.key == b"table:display" => {
                                vh.set_display(parse_bool(&attr.value)?);
                            }
                            attr if attr.key == b"table:title" => {
                                vh.set_title(Some(parse_string(&attr.value)?));
                            }
                            attr => {
                                dump_unused("read_validations", xml_tag.name(), &attr)?;
                            }
                        }
                    }
                    let txt = read_text_or_tag(bs, b"table:help-message", xml, xml_tag, empty_tag)?;
                    match txt {
                        TextContent::Empty => {}
                        TextContent::Xml(txt) => {
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
                }
                // no macros
                // b"office:event-listeners"
                // b"table:error-macro"
                _ => {
                    dump_unused2("read_validations", &evt)?;
                }
            },
            Event::End(ref e) => match e.name() {
                b"table:content-validation" => {
                    book.add_validation(valid);
                    valid = Validation::new();
                }
                // no macros
                // b"office:event-listeners"
                // b"table:error-macro"
                b"table:content-validations" => {
                    break;
                }
                _ => {}
            },
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

// read the master-styles tag
fn read_master_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(), OdsError> {
    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_master_styles {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                b"style:master-page" => {
                    read_master_page(bs, book, origin, xml, xml_tag)?;
                }
                _ => {
                    dump_unused2("read_master_styles", &evt)?;
                }
            },
            Event::Text(_) => (),
            Event::End(ref e) => {
                if e.name() == b"office:master-styles" {
                    break;
                }
            }
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
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
) -> Result<(), OdsError> {
    let mut masterpage = MasterPage::new_empty();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                masterpage.set_name(parse_string(&attr.value)?);
            }
            attr if attr.key == b"style:page-layout-name" => {
                masterpage.set_pagestyle(&parse_string(&attr.value)?.into());
            }
            attr => {
                dump_unused("read_master_page", xml_tag.name(), &attr)?;
            }
        }
    }

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_master_page {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) => match xml_tag.name() {
                b"style:header" => {
                    masterpage.set_header(read_headerfooter(bs, b"style:header", xml, xml_tag)?);
                }
                b"style:header-first" => {
                    masterpage.set_header_first(read_headerfooter(
                        bs,
                        b"style:header-first",
                        xml,
                        xml_tag,
                    )?);
                }
                b"style:header-left" => {
                    masterpage.set_header_left(read_headerfooter(
                        bs,
                        b"style:header-left",
                        xml,
                        xml_tag,
                    )?);
                }
                b"style:footer" => {
                    masterpage.set_footer(read_headerfooter(bs, b"style:footer", xml, xml_tag)?);
                }
                b"style:footer-first" => {
                    masterpage.set_footer_first(read_headerfooter(
                        bs,
                        b"style:footer-first",
                        xml,
                        xml_tag,
                    )?);
                }
                b"style:footer-left" => {
                    masterpage.set_footer_left(read_headerfooter(
                        bs,
                        b"style:footer-left",
                        xml,
                        xml_tag,
                    )?);
                }
                _ => {
                    dump_unused2("read_master_page", &evt)?;
                }
            },

            Event::Empty(_) => {
                // empty header/footer tags can be skipped.
            }

            Event::Text(_) => (),
            Event::End(ref e) => {
                if e.name() == b"style:master-page" {
                    break;
                }
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
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
) -> Result<HeaderFooter, OdsError> {
    let mut hf = HeaderFooter::new();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:display" => {
                hf.set_display(parse_bool(&attr.value)?);
            }
            attr => {
                dump_unused("read_headerfooter", xml_tag.name(), &attr)?;
            }
        }
    }

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_headerfooter {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:region-left" => {
                        let cm =
                            read_xml_content(bs, b"style:region-left", xml, xml_tag, empty_tag)?;
                        if let Some(cm) = cm {
                            hf.set_left(cm);
                        }
                    }
                    b"style:region-center" => {
                        let cm =
                            read_xml_content(bs, b"style:region-center", xml, xml_tag, empty_tag)?;
                        if let Some(cm) = cm {
                            hf.set_center(cm);
                        }
                    }
                    b"style:region-right" => {
                        let cm =
                            read_xml_content(bs, b"style:region-right", xml, xml_tag, empty_tag)?;
                        if let Some(cm) = cm {
                            hf.set_right(cm);
                        }
                    }
                    b"text:p" => {
                        // todo: in table:cell there can be multiple text:p. applies here too?
                        let cm = read_xml(bs, b"text:p", xml, xml_tag, empty_tag)?;
                        hf.set_content(cm);
                    }
                    // no other tags supported for now.
                    _ => {
                        dump_unused2("read_headerfooter", &evt)?;
                    }
                }
            }
            Event::Text(_) => (),
            Event::End(ref e) => {
                if e.name() == end_tag {
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
fn read_styles_tag(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // not attributes
) -> Result<(), OdsError> {
    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_styles_tag {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:style" => {
                        read_style_style(
                            bs,
                            book,
                            origin,
                            StyleUse::Named,
                            b"style:style",
                            xml,
                            xml_tag,
                            empty_tag,
                        )?;
                    }
                    b"style:default-style" => {
                        read_style_style(
                            bs,
                            book,
                            origin,
                            StyleUse::Default,
                            b"style:default-style",
                            xml,
                            xml_tag,
                            empty_tag,
                        )?;
                    }
                    b"number:boolean-style"
                    | b"number:date-style"
                    | b"number:time-style"
                    | b"number:number-style"
                    | b"number:currency-style"
                    | b"number:percentage-style"
                    | b"number:text-style" => {
                        read_value_format(bs, book, origin, StyleUse::Named, xml, xml_tag)?;
                    }
                    // style:default-page-layout
                    _ => {
                        dump_unused2("read_styles_tag", &evt)?;
                    }
                }
            }
            Event::Text(_) => (),
            Event::End(ref e) => {
                if e.name() == b"office:styles" {
                    break;
                }
            }
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
fn read_auto_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(), OdsError> {
    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if DUMP_XML {
            println!(" read_auto_styles {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:style" => {
                        read_style_style(
                            bs,
                            book,
                            origin,
                            StyleUse::Automatic,
                            b"style:style",
                            xml,
                            xml_tag,
                            empty_tag,
                        )?;
                    }
                    b"number:boolean-style"
                    | b"number:date-style"
                    | b"number:time-style"
                    | b"number:number-style"
                    | b"number:currency-style"
                    | b"number:percentage-style"
                    | b"number:text-style" => {
                        read_value_format(bs, book, origin, StyleUse::Automatic, xml, xml_tag)?;
                    }
                    // style:default-page-layout
                    b"style:page-layout" => {
                        read_page_style(bs, book, xml, xml_tag)?;
                    }
                    _ => {
                        dump_unused2("read_auto_styles", &evt)?;
                    }
                }
            }
            Event::Text(_) => (),
            Event::End(ref e) => {
                if e.name() == b"office:automatic-styles" {
                    break;
                }
            }
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
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
) -> Result<(), OdsError> {
    let mut valuestyle = ValueFormat::new_empty();
    valuestyle.set_origin(origin);
    valuestyle.set_styleuse(styleuse);
    // Styles with content information are stored before completion.
    let mut valuestyle_part = None;

    match xml_tag.name() {
        b"number:boolean-style" => {
            read_value_format_attr(ValueType::Boolean, &mut valuestyle, xml_tag)?
        }
        b"number:date-style" => {
            read_value_format_attr(ValueType::DateTime, &mut valuestyle, xml_tag)?
        }
        b"number:time-style" => {
            read_value_format_attr(ValueType::TimeDuration, &mut valuestyle, xml_tag)?
        }
        b"number:number-style" => {
            read_value_format_attr(ValueType::Number, &mut valuestyle, xml_tag)?
        }
        b"number:currency-style" => {
            read_value_format_attr(ValueType::Currency, &mut valuestyle, xml_tag)?
        }
        b"number:percentage-style" => {
            read_value_format_attr(ValueType::Percentage, &mut valuestyle, xml_tag)?
        }
        b"number:text-style" => read_value_format_attr(ValueType::Text, &mut valuestyle, xml_tag)?,
        _ => {
            if DUMP_UNUSED {
                let n = xml.decode(xml_tag.name())?;
                println!(" read_value_format unused {}", n);
            }
        }
    }

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_value_format {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"number:boolean" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Boolean)?)
                    }
                    b"number:number" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Number)?)
                    }
                    b"number:scientific-number" => valuestyle.push_part(read_part(
                        bs,
                        xml,
                        xml_tag,
                        FormatPartType::ScientificNumber,
                    )?),
                    b"number:day" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Day)?)
                    }
                    b"number:month" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Month)?)
                    }
                    b"number:year" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Year)?)
                    }
                    b"number:era" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Era)?)
                    }
                    b"number:day-of-week" => valuestyle.push_part(read_part(
                        bs,
                        xml,
                        xml_tag,
                        FormatPartType::DayOfWeek,
                    )?),
                    b"number:week-of-year" => valuestyle.push_part(read_part(
                        bs,
                        xml,
                        xml_tag,
                        FormatPartType::WeekOfYear,
                    )?),
                    b"number:quarter" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Quarter)?)
                    }
                    b"number:hours" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Hours)?)
                    }
                    b"number:minutes" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Minutes)?)
                    }
                    b"number:seconds" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Seconds)?)
                    }
                    b"number:fraction" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Fraction)?)
                    }
                    b"number:am-pm" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::AmPm)?)
                    }
                    b"number:text-content" => valuestyle.push_part(read_part(
                        bs,
                        xml,
                        xml_tag,
                        FormatPartType::TextContent,
                    )?),
                    b"style:text" => {
                        valuestyle.push_part(read_part(bs, xml, xml_tag, FormatPartType::Day)?)
                    }
                    b"style:map" => valuestyle.push_stylemap(read_stylemap(xml_tag)?),
                    b"number:fill-character" => {
                        valuestyle_part =
                            Some(read_part(bs, xml, xml_tag, FormatPartType::FillCharacter)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = valuestyle_part {
                                valuestyle.push_part(part);
                            }
                            valuestyle_part = None;
                        }
                    }
                    b"number:currency-symbol" => {
                        valuestyle_part =
                            Some(read_part(bs, xml, xml_tag, FormatPartType::CurrencySymbol)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = valuestyle_part {
                                valuestyle.push_part(part);
                            }
                            valuestyle_part = None;
                        }
                    }
                    b"number:text" => {
                        valuestyle_part = Some(read_part(bs, xml, xml_tag, FormatPartType::Text)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = valuestyle_part {
                                valuestyle.push_part(part);
                            }
                            valuestyle_part = None;
                        }
                    }
                    b"style:text-properties" => copy_attr2(valuestyle.textstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_value_format", &evt)?;
                    }
                }
            }
            Event::Text(ref e) => {
                if let Some(part) = &mut valuestyle_part {
                    part.set_content(e.unescape_and_decode(xml)?);
                }
            }
            Event::End(ref e) => match e.name() {
                b"number:boolean-style"
                | b"number:date-style"
                | b"number:time-style"
                | b"number:number-style"
                | b"number:currency-style"
                | b"number:percentage-style"
                | b"number:text-style" => {
                    book.add_format(valuestyle);
                    break;
                }
                b"number:currency-symbol" | b"number:text" | b"number:fill-character" => {
                    if let Some(part) = valuestyle_part {
                        valuestyle.push_part(part);
                    }
                    valuestyle_part = None;
                }
                _ => {
                    dump_unused2("read_value_format", &evt)?;
                }
            },
            Event::Eof => break,
            _ => {
                dump_unused2("read_value_format", &evt)?;
            }
        }

        buf.clear();
    }
    bs.push(buf);

    Ok(())
}

/// Copies all the attr from the tag.
fn read_value_format_attr(
    value_type: ValueType,
    valuestyle: &mut ValueFormat,
    xml_tag: &BytesStart<'_>,
) -> Result<(), OdsError> {
    valuestyle.set_value_type(value_type);
    let name = proc_style_attr(valuestyle.attrmap_mut(), xml_tag)?;
    valuestyle.set_name(name);

    Ok(())
}

fn read_part(
    bs: &mut BufStack,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    part_type: FormatPartType,
) -> Result<FormatPart, OdsError> {
    let mut part = FormatPart::new(part_type);
    copy_attr2(part.attrmap_mut(), xml_tag)?;

    // There is one relevant subtag embedded-text.
    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_part {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag2) | Event::Empty(ref xml_tag2) => match xml_tag2.name() {
                b"number:embedded-text" => {
                    for attr in xml_tag2.attributes().with_checks(false) {
                        let attr = attr?;
                        match attr.key {
                            b"number:position" => {
                                part.set_position(parse_u32(&attr.value)?);
                            }
                            attr => {
                                return Err(OdsError::Ods(format!(
                                    "embedded-text: attribute unknown {} ",
                                    from_utf8(attr)?
                                )))
                            }
                        }
                    }
                }
                _ => dump_unused2("read_value_format", &evt)?,
            },
            Event::Text(ref e) => {
                part.set_content(e.unescape_and_decode(xml)?);
            }
            Event::End(ref e) => match e.name() {
                b"number:embedded-text" => {
                    break;
                }
                _ => {
                    dump_unused2("read_value_format", &evt)?;
                }
            },
            Event::Eof => break,
            _ => {}
        }
    }

    bs.push(buf);

    Ok(part)
}

#[allow(clippy::too_many_arguments)]
// style:style tag
fn read_style_style(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:family" => {
                match attr.value.as_ref() {
                    b"table" => read_tablestyle(
                        bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag,
                    )?,
                    b"table-column" => {
                        read_colstyle(bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?
                    }
                    b"table-row" => {
                        read_rowstyle(bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?
                    }
                    b"table-cell" => read_cellstyle(
                        bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag,
                    )?,
                    b"graphic" => read_graphicstyle(
                        bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag,
                    )?,
                    b"paragraph" => read_paragraphstyle(
                        bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag,
                    )?,
                    b"text" => read_textstyle(
                        bs, book, origin, styleuse, end_tag, xml, xml_tag, empty_tag,
                    )?,
                    value => {
                        return Err(OdsError::Ods(format!(
                            "style:family unknown {} ",
                            from_utf8(value)?
                        )))
                    }
                };
            }
            attr => {
                dump_unused("read_style_style", xml_tag.name(), &attr)?;
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
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = TableStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_tablestyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_table_style {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-properties" => copy_attr2(style.tablestyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_table_style", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
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
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = RowStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_rowstyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_rowstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-row-properties" => copy_attr2(style.rowstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_rowstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
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
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = ColStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_colstyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_colstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-column-properties" => copy_attr2(style.colstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_colstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
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
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = CellStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_cellstyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_cellstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-cell-properties" => copy_attr2(style.cellstyle_mut(), xml_tag)?,
                    b"style:text-properties" => copy_attr2(style.textstyle_mut(), xml_tag)?,
                    b"style:paragraph-properties" => {
                        copy_attr2(style.paragraphstyle_mut(), xml_tag)?
                    }
                    // b"style:graphic-properties" => copy_attr(style.graphic_mut(), xml, xml_tag)?,
                    b"style:map" => style.push_stylemap(read_stylemap(xml_tag)?),

                    // b"style:tab-stops" => (),
                    // b"style:tab-stop" => {
                    //     let mut ts = TabStop::new();
                    //     copy_attr(&mut ts, xml, xml_tag)?;
                    //     style.paragraph_mut().add_tabstop(ts);
                    // }
                    _ => {
                        dump_unused2("read_cellstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_cellstyle(style);
                        break;
                    } else if e.name() == b"style:paragraph-properties" {
                        // noop
                    } else {
                        dump_unused2("read_cellstyle", &evt)?;
                    }
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
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = ParagraphStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_paragraphstyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_paragraphstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:text-properties" => copy_attr2(style.textstyle_mut(), xml_tag)?,
                    b"style:paragraph-properties" => {
                        copy_attr2(style.paragraphstyle_mut(), xml_tag)?
                    }
                    // b"style:graphic-properties" => copy_attr(style.graphic_mut(), xml, xml_tag)?,
                    // b"style:map" => style.push_stylemap(read_stylemap(xml, xml_tag)?),
                    b"style:tab-stops" => (),
                    b"style:tab-stop" => {
                        let mut ts = TabStop::new();
                        copy_attr2(ts.attrmap_mut(), xml_tag)?;
                        style.add_tabstop(ts);
                    }
                    _ => {
                        dump_unused2("read_paragraphstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_paragraphstyle(style);
                        break;
                    } else if e.name() == b"style:tab-stops"
                        || e.name() == b"style:paragraph-properties"
                    {
                        // noop
                    } else {
                        dump_unused2("read_paragraphstyle", &evt)?;
                    }
                }
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
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = TextStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_textstyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_textstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:text-properties" => copy_attr2(style.textstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_textstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_textstyle(style);
                        break;
                    } else {
                        dump_unused2("read_textstyle", &evt)?;
                    }
                }
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
fn read_graphicstyle(
    bs: &mut BufStack,
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut style = GraphicStyle::new_empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    let name = proc_style_attr(style.attrmap_mut(), xml_tag)?;
    style.set_name(name);

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_graphicstyle(style);
    } else {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_graphicstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:graphic-properties" => copy_attr2(style.graphicstyle_mut(), xml_tag)?,
                    _ => {
                        dump_unused2("read_graphicstyle", &evt)?;
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_graphicstyle(style);
                        break;
                    } else {
                        dump_unused2("read_graphicstyle", &evt)?;
                    }
                }
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

fn read_stylemap(xml_tag: &BytesStart<'_>) -> Result<StyleMap, OdsError> {
    let mut sm = StyleMap::default();
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:condition" => {
                sm.set_condition(ValueCondition::new(parse_string(&attr.value)?));
            }
            attr if attr.key == b"style:apply-style-name" => {
                sm.set_applied_style(parse_string(&attr.value)?);
            }
            attr if attr.key == b"style:base-cell-address" => {
                // todo: maybe better?
                let v = attr.unescaped_value()?;
                let mut pos = 0usize;
                sm.set_base_cell(parse_cellref(from_utf8(v.as_ref())?, &mut pos)?);
            }
            attr => {
                dump_unused("read_stylemap", xml_tag.name(), &attr)?;
            }
        }
    }

    Ok(sm)
}

/// Copies all attributes to the map, excluding "style:name" which is returned.
fn proc_style_attr<'a>(
    attrmap: &mut AttrMap2,
    xml_tag: &'a BytesStart<'_>,
) -> Result<String, OdsError> {
    let mut name = None;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                name = Some(parse_string(&attr.value)?);
            }
            attr => {
                let k = from_utf8(attr.key)?;
                let v = parse_string(&attr.value)?;
                attrmap.set_attr(k, v);
            }
        }
    }

    Ok(name.unwrap_or_default())
}

/// Copies all attributes to the given map.
fn copy_attr2(attrmap: &mut AttrMap2, xml_tag: &BytesStart<'_>) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        let attr = attr?;

        let k = from_utf8(attr.key)?;
        let v = parse_string(&attr.value)?;
        attrmap.set_attr(k, v);
    }

    Ok(())
}

fn read_styles(
    bs: &mut BufStack,
    book: &mut WorkBook,
    zip_file: &mut ZipFile<'_>,
) -> Result<(), OdsError> {
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    // Do not trim text data. All text read contains significant whitespace.
    // The rest is ignored anyway.
    //
    // xml.trim_text(true);

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_styles {:?}", evt);
        }
        match evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag) if xml_tag.name() == b"office:document-styles" => {
                // noop
            }
            Event::End(xml_tag) if xml_tag.name() == b"office:document-styles" => {
                // noop
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:font-face-decls" => {
                read_fonts(bs, book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:styles" => {
                read_styles_tag(bs, book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:automatic-styles" => {
                read_auto_styles(bs, book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:master-styles" => {
                read_master_styles(bs, book, StyleOrigin::Styles, &mut xml)?
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

fn read_settings(
    bs: &mut BufStack,
    book: &mut WorkBook,
    zip_file: &mut ZipFile<'_>,
) -> Result<(), OdsError> {
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    // Do not trim text data. All text read contains significant whitespace.
    // The rest is ignored anyway.
    //
    // xml.trim_text(true);

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_settings {:?}", evt);
        }

        match evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag) if xml_tag.name() == b"office:document-settings" => {
                // noop
            }
            Event::End(xml_tag) if xml_tag.name() == b"office:document-settings" => {
                // noop
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:settings" => {
                book.config = Detach::new(read_office_settings(bs, &mut xml)?);
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
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<Config, OdsError> {
    let mut config = Config::new();

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_office_settings {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-set" => {
                let (name, set) = read_config_item_set(bs, xml_tag, xml)?;
                config.insert(name, set);
            }
            Event::End(ref e) if e.name() == b"office:settings" => {
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

    Ok(config)
}

// read the automatic-styles tag
fn read_config_item_set(
    bs: &mut BufStack,
    xml_tag: &BytesStart<'_>,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(String, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_set = ConfigItem::new_set();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                name = Some(parse_string(&attr.value)?);
            }
            attr => {
                dump_unused("read_config_item_set", xml_tag.name(), &attr)?;
            }
        }
    }

    let name = if let Some(name) = name {
        name
    } else {
        return Err(OdsError::Ods("config-item-set without name".to_string()));
    };

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_office_item_set {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item" => {
                let (name, val) = read_config_item(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-set" => {
                let (name, val) = read_config_item_set(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-indexed" => {
                let (name, val) = read_config_item_map_indexed(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-named" => {
                let (name, val) = read_config_item_map_named(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::End(ref e) if e.name() == b"config:config-item-set" => {
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
    xml_tag: &BytesStart<'_>,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(String, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_vec = ConfigItem::new_vec();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                name = Some(parse_string(&attr.value)?);
            }
            attr => {
                dump_unused("read_config_item_map_indexed", xml_tag.name(), &attr)?;
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

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_office_item_set {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-entry" => {
                let (_, entry) = read_config_item_map_entry(bs, xml_tag, xml)?;
                config_vec.insert(index.to_string(), entry);
                index += 1;
            }
            Event::End(ref e) if e.name() == b"config:config-item-map-indexed" => {
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
    xml_tag: &BytesStart<'_>,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(String, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_map = ConfigItem::new_map();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                name = Some(parse_string(&attr.value)?);
            }
            attr => {
                dump_unused("read_config_item_map_named", xml_tag.name(), &attr)?;
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

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_config_item_map_named {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-entry" => {
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
            Event::End(ref e) if e.name() == b"config:config-item-map-named" => {
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
    xml_tag: &BytesStart<'_>,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
) -> Result<(Option<String>, ConfigItem), OdsError> {
    let mut name = None;
    let mut config_set = ConfigItem::new_entry();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                name = Some(parse_string(&attr.value)?);
            }
            attr => {
                dump_unused("read_config_item_map_entry", xml_tag.name(), &attr)?;
            }
        }
    }

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if DUMP_XML {
            println!(" read_config_item_map_entry {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item" => {
                let (name, val) = read_config_item(bs, xml_tag, xml)?;
                config_set.insert(name, ConfigItem::from(val));
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-set" => {
                let (name, val) = read_config_item_set(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-indexed" => {
                let (name, val) = read_config_item_map_indexed(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-named" => {
                let (name, val) = read_config_item_map_named(bs, xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::End(ref e) if e.name() == b"config:config-item-map-entry" => {
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
    xml_tag: &BytesStart<'_>,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    // no attributes
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

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                name = Some(parse_string(&attr.value)?);
            }
            attr if attr.key == b"config:type" => {
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
                dump_unused("read_config_item", xml_tag.name(), &attr)?;
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

    // todo: is this a good way?
    let mut value: Vec<u8> = Vec::new();

    let mut buf = bs.get_buf();
    loop {
        let evt = xml.read_event(&mut buf)?;
        match evt {
            Event::Text(ref txt) => {
                value.append(&mut Vec::from(txt.unescaped()?));
            }
            Event::End(ref e) if e.name() == b"config:config-item" => {
                let value = Cow::from(value);
                match val_type {
                    ConfigValueType::None => {}
                    ConfigValueType::Base64Binary => {
                        config_val = Some(ConfigValue::Base64Binary(parse_string(&value)?));
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

    let config_val = if let Some(config_val) = config_val {
        config_val
    } else {
        return Err(OdsError::Ods("config-item without value???".to_string()));
    };

    Ok((name, config_val))
}

// Reads a part of the XML as XmlTag's, and returns the first content XmlTag.
fn read_xml_content(
    bs: &mut BufStack,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<Option<XmlTag>, OdsError> {
    let mut xml = read_xml(bs, end_tag, xml, xml_tag, empty_tag)?;
    match xml.content().get(0) {
        None => Ok(None),
        Some(XmlContent::Tag(_)) => {
            if let XmlContent::Tag(tag) = xml.content_mut().pop().unwrap() {
                Ok(Some(tag))
            } else {
                unreachable!()
            }
        }
        Some(XmlContent::Text(_)) => Ok(None),
    }
}

// Reads a part of the XML as XmlTag's.
fn read_xml(
    bs: &mut BufStack,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<XmlTag, OdsError> {
    let mut stack = Vec::new();

    let mut tag = XmlTag::new(xml.decode(xml_tag.name())?);
    copy_attr2(tag.attrmap_mut(), xml_tag)?;
    stack.push(tag);

    if !empty_tag {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_xml {:?}", evt);
            }
            match evt {
                Event::Start(xmlbytes) => {
                    let mut tag = XmlTag::new(xml.decode(xmlbytes.name())?);
                    copy_attr2(tag.attrmap_mut(), &xmlbytes)?;
                    stack.push(tag);
                }

                Event::End(xmlbytes) => {
                    if xmlbytes.name() == end_tag {
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
                    let mut emptytag = XmlTag::new(xml.decode(xmlbytes.name())?);
                    copy_attr2(emptytag.attrmap_mut(), &xmlbytes)?;

                    if let Some(parent) = stack.last_mut() {
                        parent.add_tag(emptytag);
                    } else {
                        unreachable!()
                    }
                }

                Event::Text(xmlbytes) => {
                    if let Some(parent) = stack.last_mut() {
                        parent.add_text(xmlbytes.unescape_and_decode(xml).unwrap());
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
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile<'_>>>,
    xml_tag: &BytesStart<'_>,
    empty_tag: bool,
) -> Result<TextContent, OdsError> {
    let mut stack = Vec::new();
    let mut cellcontent = TextContent::Empty;

    // The toplevel element is passed in with the xml_tag.
    // It is only created if there are further xml tags in the
    // element. If there is only text this is not needed.
    let create_toplevel = |t: Option<String>| -> Result<XmlTag, OdsError> {
        // No parent tag on the stack. Create the parent.
        let mut toplevel = XmlTag::new(from_utf8(xml_tag.name())?);
        copy_attr2(toplevel.attrmap_mut(), xml_tag)?;
        if let Some(t) = t {
            toplevel.add_text(t);
        }
        Ok(toplevel)
    };

    if !empty_tag {
        let mut buf = bs.get_buf();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if DUMP_XML {
                println!(" read_xml {:?}", evt);
            }
            match evt {
                Event::Text(xmlbytes) => {
                    let v = xmlbytes.unescaped()?;
                    let v = from_utf8(v.as_ref())?;

                    cellcontent = match cellcontent {
                        TextContent::Empty => {
                            // Fresh plain text string.
                            TextContent::Text(v.to_string())
                        }
                        TextContent::Text(mut old_txt) => {
                            // We have a previous plain text string. Append to it.
                            old_txt.push_str(v);
                            TextContent::Text(old_txt)
                        }
                        TextContent::Xml(mut xml) => {
                            // There is already a tag. Append the text to its children.
                            xml.add_text(v);
                            TextContent::Xml(xml)
                        }
                        TextContent::XmlVec(_) => {
                            unreachable!()
                        }
                    };
                }
                Event::Start(xmlbytes) => {
                    match cellcontent {
                        TextContent::Empty => {
                            stack.push(create_toplevel(None)?);
                        }
                        TextContent::Text(old_txt) => {
                            stack.push(create_toplevel(Some(old_txt))?);
                        }
                        TextContent::Xml(parent) => {
                            stack.push(parent);
                        }
                        TextContent::XmlVec(_) => {
                            unreachable!()
                        }
                    }

                    // Set the new tag.
                    let mut new_tag = XmlTag::new(xml.decode(xmlbytes.name())?);
                    copy_attr2(new_tag.attrmap_mut(), &xmlbytes)?;
                    cellcontent = TextContent::Xml(new_tag)
                }
                Event::End(xmlbytes) => {
                    if xmlbytes.name() == end_tag {
                        if !stack.is_empty() {
                            return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(format!("XML corrupted. Endtag {} occured before all elements are closed: {:?}", from_utf8(end_tag)?, stack))));
                        }
                        break;
                    }

                    cellcontent = match cellcontent {
                        TextContent::Empty | TextContent::Text(_) => {
                            return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(format!(
                                "XML corrupted. Endtag {} occured without start tag",
                                from_utf8(xmlbytes.name())?
                            ))));
                        }
                        TextContent::Xml(tag) => {
                            if let Some(mut parent) = stack.pop() {
                                parent.add_tag(tag);
                                TextContent::Xml(parent)
                            } else {
                                return Err(OdsError::Xml(quick_xml::Error::UnexpectedToken(
                                    format!(
                                        "XML corrupted. Endtag {} occured without start tag",
                                        from_utf8(xmlbytes.name())?
                                    ),
                                )));
                            }
                        }
                        TextContent::XmlVec(_) => {
                            unreachable!()
                        }
                    }
                }
                Event::Empty(xmlbytes) => {
                    match cellcontent {
                        TextContent::Empty => {
                            stack.push(create_toplevel(None)?);
                        }
                        TextContent::Text(txt) => {
                            stack.push(create_toplevel(Some(txt))?);
                        }
                        TextContent::Xml(parent) => {
                            stack.push(parent);
                        }
                        TextContent::XmlVec(_) => {
                            unreachable!()
                        }
                    }
                    if let Some(mut parent) = stack.pop() {
                        // Create the tag and append it immediately to the parent.
                        let mut emptytag = XmlTag::new(xml.decode(xmlbytes.name())?);
                        copy_attr2(emptytag.attrmap_mut(), &xmlbytes)?;
                        parent.add_tag(emptytag);

                        cellcontent = TextContent::Xml(parent);
                    } else {
                        unreachable!()
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

fn dump_unused(func: &str, tag: &[u8], attr: &Attribute<'_>) -> Result<(), OdsError> {
    if DUMP_UNUSED {
        let tag = from_utf8(tag)?;
        let key = from_utf8(attr.key)?;
        let value = from_utf8(attr.value.as_ref())?;
        println!("unused attr: {} {} ({}:{})", func, tag, key, value);
    }
    Ok(())
}

fn dump_unused2(func: &str, evt: &Event<'_>) -> Result<(), OdsError> {
    if DUMP_UNUSED {
        println!("unused attr: {} ({:?})", func, evt);
    }
    Ok(())
}
