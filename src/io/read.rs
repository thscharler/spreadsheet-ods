use std::convert::TryFrom;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime};
use quick_xml::events::{BytesStart, Event};
use zip::read::ZipFile;
use zip::ZipArchive;

use crate::attrmap2::AttrMap2;
use crate::condition::{Condition, ValueCondition};
use crate::config::{Config, ConfigItem, ConfigItemType, ConfigValue};
use crate::ds::detach::Detach;
use crate::error::OdsError;
use crate::format::{FormatPart, FormatPartType};
use crate::refs::{parse_cellranges, parse_cellref, CellRef};
use crate::style::stylemap::StyleMap;
use crate::style::tabstop::TabStop;
use crate::style::{
    ColStyle, FontFaceDecl, GraphicStyle, HeaderFooter, MasterPage, PageStyle, ParagraphStyle,
    RowStyle, StyleOrigin, StyleUse, TableStyle, TextStyle,
};
use crate::text::{TextP, TextTag};
use crate::validation::{
    MessageType, Validation, ValidationDisplay, ValidationError, ValidationHelp,
};
use crate::xmltree::{XmlContent, XmlTag};
use crate::{
    ucell, CellStyle, ColRange, Length, RowRange, SCell, Sheet, SplitMode, Value, ValueFormat,
    ValueType, Visibility, WorkBook,
};

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
    let mut book = WorkBook::new();

    read_content(&mut book, &mut zip.by_name("content.xml")?)?;
    read_styles(&mut book, &mut zip.by_name("styles.xml")?)?;
    // may not exist.
    if let Ok(mut z) = zip.by_name("settings.xml") {
        read_settings(&mut book, &mut z)?;
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
                sheet.config_mut().cursor_x = *n as ucell;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["CursorPositionY"]) {
                sheet.config_mut().cursor_y = *n as ucell;
            }
            if let Some(ConfigValue::Short(n)) = cc.get_value_rec(&["HorizontalSplitMode"]) {
                sheet.config_mut().hor_split_mode = SplitMode::try_from(*n)?;
            }
            if let Some(ConfigValue::Short(n)) = cc.get_value_rec(&["VerticalSplitMode"]) {
                sheet.config_mut().vert_split_mode = SplitMode::try_from(*n)?;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["HorizontalSplitPosition"]) {
                sheet.config_mut().hor_split_pos = *n as ucell;
            }
            if let Some(ConfigValue::Int(n)) = cc.get_value_rec(&["VerticalSplitPosition"]) {
                sheet.config_mut().vert_split_pos = *n as ucell;
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
fn read_content(book: &mut WorkBook, zip_file: &mut ZipFile) -> Result<(), OdsError> {
    // xml parser
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    xml.trim_text(true);

    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if cfg!(feature = "dump_xml") {
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
                            let v = attr.unescape_and_decode_value(&xml)?;
                            book.set_version(v);
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
                read_fonts(book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:styles" =>
                read_styles_tag(book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:automatic-styles" =>
                read_auto_styles(book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:master-styles" =>
                read_master_styles(book, StyleOrigin::Content, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:content-validations" =>
                read_validations(book, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table" =>
                book.push_sheet(read_table(&mut xml, xml_tag)?),

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
                let v = read_xml(xml_tag.name(), &mut xml, &xml_tag, empty_tag)?;
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
                if cfg!(feature = "dump_unused") { println!(" unused read_content {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// Reads the table.
fn read_table(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: BytesStart,
) -> Result<Sheet, OdsError> {
    let mut sheet = Sheet::new();

    read_table_attr(&mut sheet, &xml, xml_tag)?;

    // Position within table-columns
    let mut table_col: ucell = 0;

    // Cell position
    let mut row: ucell = 0;
    let mut col: ucell = 0;

    // Rows can be repeated. In reality only empty ones ever are.
    let mut row_repeat: ucell = 1;
    let mut rowstyle: Option<String> = None;
    let mut row_cellstyle: Option<String> = None;
    let mut row_visible: Visibility = Default::default();

    let mut col_range_from = 0;
    let mut row_range_from = 0;

    let mut buf = Vec::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if cfg!(feature = "dump_xml") {
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
                sheet.extra.push(read_xml(xml_tag.name(), xml, &xml_tag, empty_tag)?);
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
                table_col = read_table_col_attr(&mut sheet, table_col, xml, &xml_tag)?;
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
                let (repeat, style, cellstyle, visible) = read_table_row_attr(xml, xml_tag)?;
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
                col = read_empty_table_cell(&mut sheet, row, col, xml, xml_tag)?;
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-cell" || xml_tag.name() == b"table:covered-table-cell" => {
                col = read_table_cell(&mut sheet, row, col, xml, xml_tag)?;
            }

            _ => {
                if cfg!(feature = "dump_unused") { println!(" unused read_table {:?}", evt); }
            }
        }
    }

    Ok(sheet)
}

// Reads the table attributes.
fn read_table_attr(
    sheet: &mut Sheet,
    xml: &quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: BytesStart,
) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:name" => {
                let v = attr.unescape_and_decode_value(xml)?;
                sheet.set_name(v);
            }
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(xml)?;
                sheet.set_style(&v.into());
            }
            attr if attr.key == b"table:print" => {
                let v = attr.unescape_and_decode_value(xml)?;
                sheet.set_print(v.parse()?);
            }
            attr if attr.key == b"table:display" => {
                let v = attr.unescape_and_decode_value(xml)?;
                sheet.set_display(v.parse()?);
            }
            attr if attr.key == b"table:print-ranges" => {
                let v = attr.unescape_and_decode_value(xml)?;
                let mut pos = 0usize;
                sheet.print_ranges = parse_cellranges(v.as_str(), &mut pos)?;
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_table_attr unused {} {} {}", n, k, v);
                }
            }
        }
    }

    Ok(())
}

// Reads table-row attributes. Returns the repeat-count.
fn read_table_row_attr(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: BytesStart,
) -> Result<(ucell, Option<String>, Option<String>, Visibility), OdsError> {
    let mut row_repeat: ucell = 1;
    let mut row_visible: Visibility = Default::default();
    let mut rowstyle: Option<String> = None;
    let mut row_cellstyle: Option<String> = None;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            // table:default-cell-style-name 19.615, table:visibility 19.749 and xml:id 19.914.
            attr if attr.key == b"table:number-rows-repeated" => {
                let v = attr.unescaped_value()?;
                let v = xml.decode(v.as_ref())?;
                row_repeat = v.parse::<ucell>()?;
            }
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                rowstyle = Some(v);
            }
            attr if attr.key == b"table:default-cell-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                row_cellstyle = Some(v);
            }
            attr if attr.key == b"table:visibility" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                row_visible = v.parse()?;
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_table_row_attr unused {} {} {}", n, k, v);
                }
            }
        }
    }

    Ok((row_repeat, rowstyle, row_cellstyle, row_visible))
}

// Reads the table-column attributes. Creates as many copies as indicated.
fn read_table_col_attr(
    sheet: &mut Sheet,
    mut table_col: ucell,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<ucell, OdsError> {
    let mut style = None;
    let mut cellstyle = None;
    let mut repeat: ucell = 1;
    let mut visible: Visibility = Default::default();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-columns-repeated" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                repeat = v.parse()?;
            }
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style = Some(v);
            }
            attr if attr.key == b"table:default-cell-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                cellstyle = Some(v);
            }
            attr if attr.key == b"table:visibility" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                visible = v.parse()?;
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_table_column unused {} {} {}", n, k, v);
                }
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

// Reads the cell data.
#[allow(clippy::collapsible_else_if)]
fn read_table_cell(
    sheet: &mut Sheet,
    row: ucell,
    mut col: ucell,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: BytesStart,
) -> Result<ucell, OdsError> {
    // Current cell tag
    let tag_name = xml_tag.name();

    // The current cell.
    let mut cell: SCell = SCell::new();
    // Columns can be repeated, not only empty ones.
    let mut cell_repeat: ucell = 1;
    // Decoded type.
    let mut value_type: Option<ValueType> = None;
    // Basic cell value here.
    let mut cell_value: Option<String> = None;
    // Content of the table-cell tag.
    let mut cell_content: Option<String> = None;
    // Content of the table-cell tag.
    let mut cell_content_txt: Option<Vec<TextTag>> = None;
    // Currency
    let mut cell_currency: Option<String> = None;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-columns-repeated" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                cell_repeat = v.parse::<ucell>()?;
            }
            attr if attr.key == b"table:number-rows-spanned" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                cell.span.0 = v.parse::<ucell>()?;
            }
            attr if attr.key == b"table:number-columns-spanned" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                cell.span.1 = v.parse::<ucell>()?;
            }

            attr if attr.key == b"office:value-type" => {
                value_type = match attr.unescaped_value()?.as_ref() {
                    b"string" => Some(ValueType::Text),
                    b"float" => Some(ValueType::Number),
                    b"percentage" => Some(ValueType::Percentage),
                    b"date" => Some(ValueType::DateTime),
                    b"time" => Some(ValueType::TimeDuration),
                    b"boolean" => Some(ValueType::Boolean),
                    b"currency" => Some(ValueType::Currency),
                    other => return Err(OdsError::Ods(format!("Unknown cell-type {:?}", other))),
                }
            }
            attr if attr.key == b"calcext:value-type" => {}

            attr if attr.key == b"office:date-value" => {
                cell_value = Some(attr.unescape_and_decode_value(&xml)?)
            }
            attr if attr.key == b"office:time-value" => {
                cell_value = Some(attr.unescape_and_decode_value(&xml)?)
            }
            attr if attr.key == b"office:value" => {
                cell_value = Some(attr.unescape_and_decode_value(&xml)?)
            }
            attr if attr.key == b"office:boolean-value" => {
                cell_value = Some(attr.unescape_and_decode_value(&xml)?)
            }
            attr if attr.key == b"office:string-value" => {
                cell_value = Some(attr.unescape_and_decode_value(&xml)?)
            }

            attr if attr.key == b"office:currency" => {
                cell_currency = Some(attr.unescape_and_decode_value(&xml)?)
            }

            attr if attr.key == b"table:formula" => {
                cell.formula = Some(attr.unescape_and_decode_value(&xml)?)
            }
            attr if attr.key == b"table:style-name" => {
                cell.style = Some(attr.unescape_and_decode_value(&xml)?)
            }

            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_table_cell unused {} {} {}", n, k, v);
                }
            }
        }
    }

    let mut buf = Vec::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_table_cell {:?}", evt);
        }
        match evt {
            Event::Start(xml_tag) if xml_tag.name() == b"text:p" => {
                let (str, txt) = read_text_or_tag(b"text:p", xml, &xml_tag, false)?;
                // There can be multiple text:p elements within the cell.
                if cell_content.is_some() {
                    // Have a destructured text:p from before.
                    // Wrap up and create list.
                    let mut vec = Vec::new();

                    // destructured text:p
                    let cc = cell_content.take().unwrap();
                    vec.push(TextP::new().text(cc).into_xmltag());

                    if let Some(txt) = txt {
                        vec.push(txt);
                    } else if let Some(str) = str {
                        vec.push(TextP::new().text(str).into_xmltag());
                    }

                    cell_content = None;
                    cell_content_txt = Some(vec);
                } else if cell_content_txt.is_some() {
                    let mut vec = cell_content_txt.take().unwrap();

                    if let Some(txt) = txt {
                        vec.push(txt);
                    } else if let Some(str) = str {
                        vec.push(TextP::new().text(str).into_xmltag());
                    }

                    assert_eq!(cell_content, None);
                    cell_content_txt = Some(vec);
                } else {
                    if let Some(txt) = txt {
                        cell_content_txt = Some(vec![txt]);
                    } else if let Some(str) = str {
                        cell_content = Some(str);
                    }
                }
            }

            Event::Empty(xml_tag) if xml_tag.name() == b"text:p" => {
                // noop
            }

            Event::End(xml_tag) if xml_tag.name() == tag_name => {
                cell.value = parse_value(
                    value_type,
                    cell_value,
                    cell_content,
                    cell_content_txt,
                    cell_currency,
                    row,
                    col,
                )?;

                while cell_repeat > 1 {
                    sheet.add_cell(row, col, cell.clone());
                    col += 1;
                    cell_repeat -= 1;
                }
                sheet.add_cell(row, col, cell);
                col += 1;

                break;
            }

            Event::Eof => {
                break;
            }

            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_table_cell unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(col)
}

/// Reads a table-cell from an empty XML tag.
/// There seems to be no data associated, but it can have a style and a formula.
/// And first of all we need the repeat count for the correct placement.
fn read_empty_table_cell(
    sheet: &mut Sheet,
    row: ucell,
    mut col: ucell,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: BytesStart,
) -> Result<ucell, OdsError> {
    let mut cell = None;
    // Default advance is one column.
    let mut cell_repeat = 1;
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-columns-repeated" => {
                let v = attr.unescaped_value()?;
                let v = xml.decode(v.as_ref())?;
                cell_repeat = v.parse::<ucell>()?;
            }

            attr if attr.key == b"table:formula" => {
                cell.get_or_insert_with(SCell::new)
                    .set_formula(attr.unescape_and_decode_value(&xml)?);
            }
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                cell.get_or_insert_with(SCell::new).set_style(&v.into());
            }
            attr if attr.key == b"table:number-rows-spanned" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                let span = v.parse::<ucell>()?;

                cell.get_or_insert_with(SCell::new).set_row_span(span);
            }
            attr if attr.key == b"table:number-columns-spanned" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                let span = v.parse::<ucell>()?;

                cell.get_or_insert_with(SCell::new).set_col_span(span);
            }

            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_empty_table_cell unused {} {} {}", n, k, v);
                }
            }
        }
    }

    if let Some(cell) = cell {
        while cell_repeat > 1 {
            sheet.add_cell(row, col, cell.clone());
            col += 1;
            cell_repeat -= 1;
        }
        sheet.add_cell(row, col, cell);
        col += 1;
    } else {
        col += cell_repeat;
    }

    Ok(col)
}

// Takes a bunch of strings and converts it to something useable.
fn parse_value(
    value_type: Option<ValueType>,
    cell_value: Option<String>,
    cell_content: Option<String>,
    cell_content_txt: Option<Vec<TextTag>>,
    cell_currency: Option<String>,
    row: ucell,
    col: ucell,
) -> Result<Value, OdsError> {
    if let Some(value_type) = value_type {
        match value_type {
            ValueType::Empty => Ok(Value::Empty),
            ValueType::Text => {
                if let Some(cell_value) = cell_value {
                    Ok(Value::Text(cell_value))
                } else if let Some(cell_content_txt) = cell_content_txt {
                    Ok(Value::TextXml(cell_content_txt))
                } else if let Some(cell_content) = cell_content {
                    Ok(Value::Text(cell_content))
                } else {
                    Ok(Value::Text("".to_string()))
                }
            }
            ValueType::TextXml => unreachable!(),
            ValueType::Number => {
                if let Some(cell_value) = cell_value {
                    let f = cell_value.parse::<f64>()?;
                    Ok(Value::Number(f))
                } else {
                    Err(OdsError::Ods(format!(
                        "{} has type number, but no value!",
                        CellRef::local(row, col)
                    )))
                }
            }
            ValueType::DateTime => {
                if let Some(cell_value) = cell_value {
                    let dt = if cell_value.len() == 10 {
                        NaiveDate::parse_from_str(cell_value.as_str(), "%Y-%m-%d")?.and_hms(0, 0, 0)
                    } else {
                        NaiveDateTime::parse_from_str(cell_value.as_str(), "%Y-%m-%dT%H:%M:%S%.f")?
                    };

                    Ok(Value::DateTime(dt))
                } else {
                    Err(OdsError::Ods(format!(
                        "{} has type datetime, but no value!",
                        CellRef::local(row, col)
                    )))
                }
            }
            ValueType::TimeDuration => {
                if let Some(mut cell_value) = cell_value {
                    let mut hour: i32 = 0;
                    let mut have_hour = false;
                    let mut min: i32 = 0;
                    let mut have_min = false;
                    let mut sec: i32 = 0;
                    let mut have_sec = false;
                    let mut nanos: i64 = 0;
                    let mut nanos_digits: u8 = 0;

                    for c in cell_value.drain(..) {
                        match c {
                            'P' | 'T' => {}
                            '0'..='9' => {
                                if !have_hour {
                                    hour = hour * 10 + (c as i32 - '0' as i32);
                                } else if !have_min {
                                    min = min * 10 + (c as i32 - '0' as i32);
                                } else if !have_sec {
                                    sec = sec * 10 + (c as i32 - '0' as i32);
                                } else {
                                    nanos = nanos * 10 + (c as i64 - '0' as i64);
                                    nanos_digits += 1;
                                }
                            }
                            'H' => have_hour = true,
                            'M' => have_min = true,
                            '.' => have_sec = true,
                            'S' => {}
                            _ => {}
                        }
                    }
                    // unseen nano digits
                    while nanos_digits < 9 {
                        nanos *= 10;
                        nanos_digits += 1;
                    }

                    let secs = (hour * 3600 + min * 60 + sec) as i64;
                    let dur = Duration::seconds(secs) + Duration::nanoseconds(nanos);

                    Ok(Value::TimeDuration(dur))
                } else {
                    Err(OdsError::Ods(format!(
                        "{} has type time-duration, but no value!",
                        CellRef::local(row, col)
                    )))
                }
            }
            ValueType::Boolean => {
                if let Some(cell_value) = cell_value {
                    Ok(Value::Boolean(&cell_value == "true"))
                } else {
                    Err(OdsError::Ods(format!(
                        "{} has type boolean, but no value!",
                        CellRef::local(row, col)
                    )))
                }
            }
            ValueType::Currency => {
                if let Some(cell_value) = cell_value {
                    let f = cell_value.parse::<f64>()?;
                    if let Some(cell_currency) = cell_currency {
                        Ok(Value::Currency(cell_currency, f))
                    } else {
                        Err(OdsError::Ods(format!(
                            "{} has type currency, but no value!",
                            CellRef::local(row, col)
                        )))
                    }
                } else {
                    Err(OdsError::Ods(format!(
                        "{} has type currency, but no value!",
                        CellRef::local(row, col)
                    )))
                }
            }
            ValueType::Percentage => {
                if let Some(cell_value) = cell_value {
                    let f = cell_value.parse::<f64>()?;
                    Ok(Value::Percentage(f))
                } else {
                    Err(OdsError::Ods(format!(
                        "{} has type percentage, but no value!",
                        CellRef::local(row, col)
                    )))
                }
            }
        }
    } else {
        // could be an image or whatever
        Ok(Value::Empty)
    }
}

// reads a font-face
fn read_fonts(
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut font: FontFaceDecl = FontFaceDecl::new();
    font.set_origin(origin);

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_fonts {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                b"style:font-face" => {
                    for attr in xml_tag.attributes().with_checks(false) {
                        match attr? {
                            attr if attr.key == b"style:name" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                font.set_name(v);
                            }
                            attr => {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                font.attrmap_mut().set_attr(k, v);
                            }
                        }
                    }

                    book.add_font(font);
                    font = FontFaceDecl::new();
                    font.set_origin(StyleOrigin::Content);
                }
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_fonts unused {:?}", evt);
                    }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_fonts unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

// reads the page-layout tag
fn read_page_style(
    book: &mut WorkBook,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut pl = PageStyle::new("");
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                pl.set_name(v);
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_page_layout unused {} {} {}", n, k, v);
                }
            }
        }
    }

    let mut headerstyle = false;
    let mut footerstyle = false;

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_page_layout {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:page-layout-properties" => copy_attr2(pl.style_mut(), xml, xml_tag)?,
                    b"style:header-style" => headerstyle = true,
                    b"style:footer-style" => footerstyle = true,
                    b"style:header-footer-properties" => {
                        if headerstyle {
                            copy_attr2(pl.headerstyle_mut().style_mut(), xml, xml_tag)?;
                        }
                        if footerstyle {
                            copy_attr2(pl.footerstyle_mut().style_mut(), xml, xml_tag)?;
                        }
                    }
                    b"style:background-image" => {
                        // noop for now. sets the background transparent.
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_page_layout unused {:?}", evt);
                        }
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
                    if cfg!(feature = "dump_unused") {
                        println!(" read_page_layout unused {:?}", evt);
                    }
                }
            },
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_page_layout unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    book.add_pagestyle(pl);

    Ok(())
}

fn read_validations(
    book: &mut WorkBook,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut valid = Validation::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if cfg!(feature = "dump_xml") {
            println!(" read_master_styles {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                b"table:content-validation" => {
                    for attr in xml_tag.attributes().with_checks(false) {
                        match attr? {
                            attr if attr.key == b"table:name" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                valid.set_name(v);
                            }
                            attr if attr.key == b"table:condition" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                // split off 'of:' prefix

                                valid.set_condition(Condition::new(v.split_at(3).1));
                            }
                            attr if attr.key == b"table:allow-empty-cell" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                valid.set_allow_empty(v.parse()?);
                            }
                            attr if attr.key == b"table:base-cell-address" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                valid.set_base_cell(CellRef::try_from(v.as_str())?);
                            }
                            attr if attr.key == b"table:display-list" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                valid.set_display(ValidationDisplay::try_from(v.as_str())?);
                            }
                            attr => {
                                if cfg!(feature = "dump_unused") {
                                    println!(" read_validations unused attr {:?}", attr);
                                }
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
                                let v = attr.unescape_and_decode_value(&xml)?;
                                ve.set_display(v.parse()?);
                            }
                            attr if attr.key == b"table:message-type" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                let mt = match v.as_str() {
                                    "stop" => MessageType::Error,
                                    "warning" => MessageType::Warning,
                                    "information" => MessageType::Info,
                                    _ => {
                                        return Err(OdsError::Parse(format!(
                                            "unknown message-type {}",
                                            v
                                        )))
                                    }
                                };
                                ve.set_msg_type(mt);
                            }
                            attr if attr.key == b"table:title" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                ve.set_title(Some(v));
                            }
                            attr => {
                                if cfg!(feature = "dump_unused") {
                                    println!(" read_validations unused attr {:?}", attr);
                                }
                            }
                        }
                    }
                    let (_str, txt) =
                        read_text_or_tag(b"table:error-message", xml, &xml_tag, empty_tag)?;
                    ve.set_text(txt);

                    valid.set_err(Some(ve));
                }
                b"table:help-message" => {
                    let mut vh = ValidationHelp::new();

                    for attr in xml_tag.attributes().with_checks(false) {
                        match attr? {
                            attr if attr.key == b"table:display" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                vh.set_display(v.parse()?);
                            }
                            attr if attr.key == b"table:title" => {
                                let v = attr.unescape_and_decode_value(&xml)?;
                                vh.set_title(Some(v));
                            }
                            attr => {
                                if cfg!(feature = "dump_unused") {
                                    println!(" read_validations unused attr {:?}", attr);
                                }
                            }
                        }
                    }
                    let (_str, txt) =
                        read_text_or_tag(b"table:help-message", xml, &xml_tag, empty_tag)?;
                    vh.set_text(txt);

                    valid.set_help(Some(vh));
                }
                // no macros
                // b"office:event-listeners"
                // b"table:error-macro"
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_validations unused {:?}", evt);
                    }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_validations unused {:?}", evt);
                }
            }
        }
    }
    Ok(())
}

// read the master-styles tag
fn read_master_styles(
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_master_styles {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                b"style:master-page" => {
                    read_master_page(book, origin, xml, xml_tag)?;
                }
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_master_styles unused {:?}", evt);
                    }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_master_styles unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

// read the master-page tag
fn read_master_page(
    book: &mut WorkBook,
    _origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut masterpage = MasterPage::empty();
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                let name = attr.unescape_and_decode_value(&xml)?;
                masterpage.set_name(name);
            }
            attr if attr.key == b"style:page-layout-name" => {
                let name = attr.unescape_and_decode_value(&xml)?;
                masterpage.set_pagestyle(&name.into());
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_master_page unused {} {} {}", n, k, v);
                }
            }
        }
    }

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_master_page {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) => match xml_tag.name() {
                b"style:header" => {
                    masterpage.set_header(read_headerfooter(b"style:header", xml, xml_tag)?);
                }
                b"style:header-first" => {
                    masterpage.set_header_first(read_headerfooter(
                        b"style:header-first",
                        xml,
                        xml_tag,
                    )?);
                }
                b"style:header-left" => {
                    masterpage.set_header_left(read_headerfooter(
                        b"style:header-left",
                        xml,
                        xml_tag,
                    )?);
                }
                b"style:footer" => {
                    masterpage.set_footer(read_headerfooter(b"style:footer", xml, xml_tag)?);
                }
                b"style:footer-first" => {
                    masterpage.set_footer_first(read_headerfooter(
                        b"style:footer-first",
                        xml,
                        xml_tag,
                    )?);
                }
                b"style:footer-left" => {
                    masterpage.set_footer_left(read_headerfooter(
                        b"style:footer-left",
                        xml,
                        xml_tag,
                    )?);
                }
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_master_page unused {:?}", evt);
                    }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_master_page unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    book.add_masterpage(masterpage);

    Ok(())
}

// reads any header or footer tags
fn read_headerfooter(
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<HeaderFooter, OdsError> {
    let mut buf = Vec::new();

    let mut hf = HeaderFooter::new();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:display" => {
                let display = attr.unescape_and_decode_value(&xml)?;
                hf.set_display(display == "true");
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_headerfooter unused {} {} {}", n, k, v);
                }
            }
        }
    }

    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if cfg!(feature = "dump_xml") {
            println!(" read_headerfooter {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:region-left" => {
                        let cm = read_xml_content(b"style:region-left", xml, &xml_tag, empty_tag)?;
                        if let Some(cm) = cm {
                            hf.set_left(cm);
                        }
                    }
                    b"style:region-center" => {
                        let cm =
                            read_xml_content(b"style:region-center", xml, &xml_tag, empty_tag)?;
                        if let Some(cm) = cm {
                            hf.set_center(cm);
                        }
                    }
                    b"style:region-right" => {
                        let cm = read_xml_content(b"style:region-right", xml, &xml_tag, empty_tag)?;
                        if let Some(cm) = cm {
                            hf.set_right(cm);
                        }
                    }
                    b"text:p" => {
                        // todo: in table:cell there can be multiple text:p. applies here too?
                        let cm = read_xml(b"text:p", xml, &xml_tag, empty_tag)?;
                        hf.set_content(cm);
                    }
                    // no other tags supported for now.
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_headerfooter unused {:?}", evt);
                        }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_headerfooter unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(hf)
}

// reads the office-styles tag
fn read_styles_tag(
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // not attributes
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if cfg!(feature = "dump_xml") {
            println!(" read_styles_tag {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:style" => {
                        read_style_style(
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
                        read_value_format(book, origin, StyleUse::Named, xml, xml_tag)?;
                    }
                    // style:default-page-layout
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_styles_tag unused {:?}", evt);
                        }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_styles_tag unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

// read the automatic-styles tag
fn read_auto_styles(
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = matches!(evt, Event::Empty(_));
        if cfg!(feature = "dump_xml") {
            println!(" read_auto_styles {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:style" => {
                        read_style_style(
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
                        read_value_format(book, origin, StyleUse::Automatic, xml, xml_tag)?;
                    }
                    // style:default-page-layout
                    b"style:page-layout" => {
                        read_page_style(book, xml, xml_tag)?;
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_auto_styles unused {:?}", evt);
                        }
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
                if cfg!(feature = "dump_unused") {
                    println!(" read_auto_styles unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

// Reads any of the number:xxx tags
fn read_value_format(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut valuestyle = ValueFormat::new();
    valuestyle.set_origin(origin);
    valuestyle.set_styleuse(styleuse);
    // Styles with content information are stored before completion.
    let mut valuestyle_part = None;

    match xml_tag.name() {
        b"number:boolean-style" => {
            read_value_format_attr(ValueType::Boolean, &mut valuestyle, xml, xml_tag)?
        }
        b"number:date-style" => {
            read_value_format_attr(ValueType::DateTime, &mut valuestyle, xml, xml_tag)?
        }
        b"number:time-style" => {
            read_value_format_attr(ValueType::TimeDuration, &mut valuestyle, xml, xml_tag)?
        }
        b"number:number-style" => {
            read_value_format_attr(ValueType::Number, &mut valuestyle, xml, xml_tag)?
        }
        b"number:currency-style" => {
            read_value_format_attr(ValueType::Currency, &mut valuestyle, xml, xml_tag)?
        }
        b"number:percentage-style" => {
            read_value_format_attr(ValueType::Percentage, &mut valuestyle, xml, xml_tag)?
        }
        b"number:text-style" => {
            read_value_format_attr(ValueType::Text, &mut valuestyle, xml, xml_tag)?
        }
        _ => {
            if cfg!(feature = "dump_unused") {
                let n = xml.decode(xml_tag.name())?;
                println!(" read_value_format unused {}", n);
            }
        }
    }

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_value_format {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"number:boolean" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Boolean)?)
                    }
                    b"number:number" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Number)?)
                    }
                    b"number:scientific-number" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Scientific)?)
                    }
                    b"number:day" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Day)?)
                    }
                    b"number:month" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Month)?)
                    }
                    b"number:year" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Year)?)
                    }
                    b"number:era" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Era)?)
                    }
                    b"number:day-of-week" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::DayOfWeek)?)
                    }
                    b"number:week-of-year" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::WeekOfYear)?)
                    }
                    b"number:quarter" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Quarter)?)
                    }
                    b"number:hours" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Hours)?)
                    }
                    b"number:minutes" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Minutes)?)
                    }
                    b"number:seconds" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Seconds)?)
                    }
                    b"number:fraction" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Fraction)?)
                    }
                    b"number:am-pm" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::AmPm)?)
                    }
                    b"number:embedded-text" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::EmbeddedText)?)
                    }
                    b"number:text-content" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::TextContent)?)
                    }
                    b"style:text" => {
                        valuestyle.push_part(read_part(xml, xml_tag, FormatPartType::Day)?)
                    }
                    b"style:map" => valuestyle.push_stylemap(read_stylemap(xml, xml_tag)?),
                    b"number:currency-symbol" => {
                        valuestyle_part =
                            Some(read_part(xml, xml_tag, FormatPartType::CurrencySymbol)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = valuestyle_part {
                                valuestyle.push_part(part);
                            }
                            valuestyle_part = None;
                        }
                    }
                    b"number:text" => {
                        valuestyle_part = Some(read_part(xml, xml_tag, FormatPartType::Text)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = valuestyle_part {
                                valuestyle.push_part(part);
                            }
                            valuestyle_part = None;
                        }
                    }
                    b"style:text-properties" => {
                        copy_attr2(valuestyle.textstyle_mut(), xml, xml_tag)?
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_value_format unused {:?}", evt);
                        }
                    }
                }
            }
            Event::Text(ref e) => {
                if let Some(part) = &mut valuestyle_part {
                    part.set_content(e.unescape_and_decode(&xml)?);
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
                b"number:currency-symbol" | b"number:text" => {
                    if let Some(part) = valuestyle_part {
                        valuestyle.push_part(part);
                    }
                    valuestyle_part = None;
                }
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_value_format unused {:?}", evt);
                    }
                }
            },
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_value_format unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

/// Copies all the attr from the tag.
fn read_value_format_attr(
    value_type: ValueType,
    valuestyle: &mut ValueFormat,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<(), OdsError> {
    valuestyle.set_value_type(value_type);

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                valuestyle.set_name(v);
            }
            attr => {
                let k = xml.decode(&attr.key)?;
                let v = attr.unescape_and_decode_value(&xml)?;
                valuestyle.attrmap_mut().set_attr(k, v);
            }
        }
    }

    Ok(())
}

fn read_part(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    part_type: FormatPartType,
) -> Result<FormatPart, OdsError> {
    let mut part = FormatPart::new(part_type);
    copy_attr2(part.attrmap_mut(), xml, xml_tag)?;
    Ok(part)
}

// style:style tag
fn read_style_style(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    match read_family_attr(xml, xml_tag)?.as_str() {
        "table" => {
            read_tablestyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        "table-row" => {
            read_rowstyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        "table-column" => {
            read_colstyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        "table-cell" => {
            read_cellstyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        "paragraph" => {
            read_paragraphstyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        "graphic" => {
            read_graphicstyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        "text" => {
            read_textstyle(book, origin, styleuse, end_tag, xml, xml_tag, empty_tag)?;
        }
        v => {
            if cfg!(feature = "dump_unused") {
                println!(" read_family_attr unused {:?}", v);
            }
        }
    }
    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_tablestyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = TableStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_tablestyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_table_style {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-properties" => copy_attr2(style.tablestyle_mut(), xml, xml_tag)?,
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_table_style unused {:?}", evt);
                        }
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_tablestyle(style);
                        break;
                    } else {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_table_style unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_table_style unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_rowstyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = RowStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_rowstyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_rowstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-row-properties" => {
                        copy_attr2(style.rowstyle_mut(), xml, xml_tag)?
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_rowstyle unused {:?}", evt);
                        }
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_rowstyle(style);
                        break;
                    } else {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_rowstyle unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_rowstyle unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_colstyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = ColStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_colstyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_colstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-column-properties" => {
                        copy_attr2(style.colstyle_mut(), xml, xml_tag)?
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_colstyle unused {:?}", evt);
                        }
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_colstyle(style);
                        break;
                    } else {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_colstyle unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_colstyle unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_cellstyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = CellStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_cellstyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_cellstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:table-cell-properties" => {
                        copy_attr2(style.cellstyle_mut(), xml, xml_tag)?
                    }
                    b"style:text-properties" => copy_attr2(style.textstyle_mut(), xml, xml_tag)?,
                    b"style:paragraph-properties" => {
                        copy_attr2(style.paragraphstyle_mut(), xml, xml_tag)?
                    }
                    // b"style:graphic-properties" => copy_attr(style.graphic_mut(), xml, xml_tag)?,
                    b"style:map" => style.push_stylemap(read_stylemap(xml, xml_tag)?),

                    // b"style:tab-stops" => (),
                    // b"style:tab-stop" => {
                    //     let mut ts = TabStop::new();
                    //     copy_attr(&mut ts, xml, xml_tag)?;
                    //     style.paragraph_mut().add_tabstop(ts);
                    // }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_cellstyle unused {:?}", evt);
                        }
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
                        if cfg!(feature = "dump_unused") {
                            println!(" read_cellstyle unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_cellstyle unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_paragraphstyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = ParagraphStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_paragraphstyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_paragraphstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:text-properties" => copy_attr2(style.textstyle_mut(), xml, xml_tag)?,
                    b"style:paragraph-properties" => {
                        copy_attr2(style.paragraphstyle_mut(), xml, xml_tag)?
                    }
                    // b"style:graphic-properties" => copy_attr(style.graphic_mut(), xml, xml_tag)?,
                    // b"style:map" => style.push_stylemap(read_stylemap(xml, xml_tag)?),
                    b"style:tab-stops" => (),
                    b"style:tab-stop" => {
                        let mut ts = TabStop::new();
                        copy_attr2(ts.attrmap_mut(), xml, xml_tag)?;
                        style.add_tabstop(ts);
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_paragraphstyle unused {:?}", evt);
                        }
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
                        if cfg!(feature = "dump_unused") {
                            println!(" read_paragraphstyle unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_paragraphstyle unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_textstyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = TextStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_textstyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_textstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:text-properties" => copy_attr2(style.textstyle_mut(), xml, xml_tag)?,
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_textstyle unused {:?}", evt);
                        }
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_textstyle(style);
                        break;
                    } else {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_textstyle unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_textstyle unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

// style:style tag
#[allow(clippy::collapsible_else_if)]
fn read_graphicstyle(
    book: &mut WorkBook,
    origin: StyleOrigin,
    styleuse: StyleUse,
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style = GraphicStyle::empty();
    style.set_origin(origin);
    style.set_styleuse(styleuse);
    style.set_name(style_name(xml, xml_tag)?);
    copy_style_attr(style.attrmap_mut(), xml, xml_tag)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_graphicstyle(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_graphicstyle {:?}", evt);
            }
            match evt {
                Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => match xml_tag.name() {
                    b"style:graphic-properties" => {
                        copy_attr2(style.graphicstyle_mut(), xml, xml_tag)?
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_graphicstyle unused {:?}", evt);
                        }
                    }
                },
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_graphicstyle(style);
                        break;
                    } else {
                        if cfg!(feature = "dump_unused") {
                            println!(" read_graphicstyle unused {:?}", evt);
                        }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_graphicstyle unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok(())
}

fn read_stylemap(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<StyleMap, OdsError> {
    let mut sm = StyleMap::default();
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:condition" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                sm.set_condition(ValueCondition::new(v));
            }
            attr if attr.key == b"style:apply-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                sm.set_applied_style(v);
            }
            attr if attr.key == b"style:base-cell-address" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                let mut pos = 0usize;
                sm.set_base_cell(parse_cellref(v.as_str(), &mut pos)?);
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_stylemap unused {} {} {}", n, k, v);
                }
            }
        }
    }

    Ok(sm)
}

fn read_family_attr(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<String, OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:family" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                return match v.as_ref() {
                    "table" | "table-column" | "table-row" | "table-cell" | "graphic"
                    | "paragraph" | "text" => Ok(v.as_str().to_string()),
                    _ => Err(OdsError::Ods(format!("style:family unknown {} ", v))),
                };
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_family_attr unused {} {} {}", n, k, v);
                }
            }
        }
    }

    Err(OdsError::Ods("no style:family".to_string()))
}

/// extract the "style:name" attr
fn style_name(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<String, OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => return Ok(attr.unescape_and_decode_value(&xml)?),
            _ => {}
        }
    }

    Ok("".to_string())
}

/// Copies all attributes to the given map, excluding "style:name"
fn copy_style_attr(
    attrmap: &mut AttrMap2,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {}
            attr => {
                let k = xml.decode(&attr.key)?;
                let v = attr.unescape_and_decode_value(&xml)?;
                attrmap.set_attr(k, v);
            }
        }
    }

    Ok(())
}

/// Copies all attributes to the given map.
fn copy_attr2(
    attrmap: &mut AttrMap2,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false).flatten() {
        let k = xml.decode(&attr.key)?;
        let v = attr.unescape_and_decode_value(&xml)?;
        attrmap.set_attr(k, v);
    }

    Ok(())
}

fn read_styles(book: &mut WorkBook, zip_file: &mut ZipFile) -> Result<(), OdsError> {
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    xml.trim_text(true);

    let mut buf = Vec::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
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
                read_fonts(book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:styles" => {
                read_styles_tag(book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:automatic-styles" => {
                read_auto_styles(book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Start(xml_tag) if xml_tag.name() == b"office:master-styles" => {
                read_master_styles(book, StyleOrigin::Styles, &mut xml)?
            }

            Event::Eof => {
                break;
            }
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_styles unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

#[allow(unused_variables)]
pub fn default_settings() -> Detach<Config> {
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

fn read_settings(book: &mut WorkBook, zip_file: &mut ZipFile) -> Result<(), OdsError> {
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    xml.trim_text(true);

    let mut buf = Vec::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
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
                book.config = Detach::new(read_office_settings(&mut xml)?);
            }

            Event::Eof => {
                break;
            }
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_settings unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(())
}

// read the automatic-styles tag
fn read_office_settings(
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<Config, OdsError> {
    let mut buf = Vec::new();

    let mut config = Config::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_office_settings {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-set" => {
                let (name, set) = read_config_item_set(xml_tag, xml)?;
                config.insert(name, set);
            }
            Event::End(ref e) if e.name() == b"office:settings" => {
                break;
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_auto_styles unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok(config)
}

// read the automatic-styles tag
fn read_config_item_set(
    xml_tag: &BytesStart,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(String, ConfigItem), OdsError> {
    let mut buf = Vec::new();

    let mut name = None;
    let mut config_set = ConfigItem::new_set();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                name = Some(v);
            }
            _ => {
                // noop
            }
        }
    }

    let name = if let Some(name) = name {
        name
    } else {
        return Err(OdsError::Ods("config-item-set without name".to_string()));
    };

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_office_item_set {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item" => {
                let (name, val) = read_config_item(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-set" => {
                let (name, val) = read_config_item_set(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-indexed" => {
                let (name, val) = read_config_item_map_indexed(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-named" => {
                let (name, val) = read_config_item_map_named(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::End(ref e) if e.name() == b"config:config-item-set" => {
                break;
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_office_item_set unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok((name, config_set))
}

// read the automatic-styles tag
fn read_config_item_map_indexed(
    xml_tag: &BytesStart,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(String, ConfigItem), OdsError> {
    let mut buf = Vec::new();

    let mut name = None;
    let mut config_vec = ConfigItem::new_vec();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                name = Some(v);
            }
            _ => {
                // noop
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

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_office_item_set {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-entry" => {
                let (_, entry) = read_config_item_map_entry(xml_tag, xml)?;
                config_vec.insert(index.to_string(), entry);
                index += 1;
            }
            Event::End(ref e) if e.name() == b"config:config-item-map-indexed" => {
                break;
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_office_item_set unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok((name, config_vec))
}

// read the automatic-styles tag
fn read_config_item_map_named(
    xml_tag: &BytesStart,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(String, ConfigItem), OdsError> {
    let mut buf = Vec::new();

    let mut name = None;
    let mut config_map = ConfigItem::new_map();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                name = Some(v);
            }
            _ => {
                // noop
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

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_config_item_map_named {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-entry" => {
                let (name, entry) = read_config_item_map_entry(xml_tag, xml)?;

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
                if cfg!(feature = "dump_unused") {
                    println!(" read_config_item_map_named unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok((name, config_map))
}

// read the automatic-styles tag
fn read_config_item_map_entry(
    xml_tag: &BytesStart,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(Option<String>, ConfigItem), OdsError> {
    let mut buf = Vec::new();

    let mut name = None;
    let mut config_set = ConfigItem::new_entry();

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                name = Some(v);
            }
            _ => {
                // noop
            }
        }
    }

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") {
            println!(" read_config_item_map_entry {:?}", evt);
        }
        match evt {
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item" => {
                let (name, val) = read_config_item(xml_tag, xml)?;
                config_set.insert(name, ConfigItem::from(val));
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-set" => {
                let (name, val) = read_config_item_set(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-indexed" => {
                let (name, val) = read_config_item_map_indexed(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::Start(ref xml_tag) if xml_tag.name() == b"config:config-item-map-named" => {
                let (name, val) = read_config_item_map_named(xml_tag, xml)?;
                config_set.insert(name, val);
            }
            Event::End(ref e) if e.name() == b"config:config-item-map-entry" => {
                break;
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_config_item_map_entry unused {:?}", evt);
                }
            }
        }

        buf.clear();
    }

    Ok((name, config_set))
}

// read the automatic-styles tag
fn read_config_item(
    xml_tag: &BytesStart,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    // no attributes
) -> Result<(String, ConfigValue), OdsError> {
    let mut buf = Vec::new();

    let mut name = None;
    let mut val_type = None;
    let mut config_val = None;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"config:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                name = Some(v);
            }
            attr if attr.key == b"config:type" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                val_type = Some(v);
            }
            _ => {
                // noop
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

    let valtype = if let Some(val_type) = val_type {
        val_type
    } else {
        return Err(OdsError::Ods(
            "config value without config:type".to_string(),
        ));
    };

    let mut value = String::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        match evt {
            Event::Text(ref txt) => {
                let txt = txt.unescape_and_decode(xml)?;
                value.push_str(txt.as_str());
            }
            Event::End(ref e) if e.name() == b"config:config-item" => {
                match valtype.as_str() {
                    "base64Binary" => {
                        config_val = Some(ConfigValue::Base64Binary(value));
                    }
                    "boolean" => {
                        let f = value.parse::<bool>()?;
                        config_val = Some(ConfigValue::Boolean(f));
                    }
                    "datetime" => {
                        let dt = if value.len() == 10 {
                            NaiveDate::parse_from_str(value.as_str(), "%Y-%m-%d")?.and_hms(0, 0, 0)
                        } else {
                            NaiveDateTime::parse_from_str(value.as_str(), "%Y-%m-%dT%H:%M:%S%.f")?
                        };
                        config_val = Some(ConfigValue::DateTime(dt));
                    }
                    "double" => {
                        let f = value.parse::<f64>()?;
                        config_val = Some(ConfigValue::Double(f));
                    }
                    "int" => {
                        let f = value.parse::<i32>()?;
                        config_val = Some(ConfigValue::Int(f));
                    }
                    "long" => {
                        let f = value.parse::<i64>()?;
                        config_val = Some(ConfigValue::Long(f));
                    }
                    "short" => {
                        let f = value.parse::<i16>()?;
                        config_val = Some(ConfigValue::Short(f));
                    }
                    "string" => {
                        config_val = Some(ConfigValue::String(value));
                    }
                    x => {
                        return Err(OdsError::Ods(format!("unknown config:type {}", x)));
                    }
                }
                break;
            }
            Event::Eof => {
                break;
            }
            _ => {
                if cfg!(feature = "dump_unused") {
                    println!(" read_config_item unused {:?}", evt);
                }
            }
        }

        if cfg!(feature = "dump_xml") {
            println!(" read_config_item {:?}", evt);
        }
        buf.clear();
    }

    let config_val = if let Some(config_val) = config_val {
        config_val
    } else {
        return Err(OdsError::Ods("config-item without value???".to_string()));
    };

    Ok((name, config_val))
}

// Reads a part of the XML as XmlTag's, and returns the first content XmlTag.
fn read_xml_content(
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<Option<XmlTag>, OdsError> {
    let mut xml = read_xml(end_tag, xml, xml_tag, empty_tag)?;
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
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<XmlTag, OdsError> {
    let mut stack = Vec::new();

    let mut tag = XmlTag::new(xml.decode(xml_tag.name())?);
    copy_attr2(tag.attrmap_mut(), xml, xml_tag)?;
    stack.push(tag);

    if !empty_tag {
        let mut buf = Vec::new();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_xml {:?}", evt);
            }
            match evt {
                Event::Start(xmlbytes) => {
                    let mut tag = XmlTag::new(xml.decode(xmlbytes.name())?);
                    copy_attr2(tag.attrmap_mut(), xml, &xmlbytes)?;
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
                    copy_attr2(emptytag.attrmap_mut(), xml, &xmlbytes)?;

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
                    if cfg!(feature = "dump_unused") {
                        println!(" read_xml unused {:?}", evt);
                    }
                }
            }
        }
    }

    assert_eq!(stack.len(), 1);
    Ok(stack.pop().unwrap())
}

// reads all the tags up to end_tag and creates a TextVec.
// if there are no tags the result is a plain String.
fn read_text_or_tag(
    end_tag: &[u8],
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
    xml_tag: &BytesStart,
    empty_tag: bool,
) -> Result<(Option<String>, Option<TextTag>), OdsError> {
    let mut str: Option<String> = None;
    let mut text: Option<TextTag> = None;

    let mut stack = Vec::<XmlTag>::new();

    if !empty_tag {
        let mut buf = Vec::new();
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") {
                println!(" read_xml {:?}", evt);
            }
            match evt {
                Event::Text(xmlbytes) => {
                    let text = xmlbytes.unescape_and_decode(xml)?;
                    if !stack.is_empty() {
                        // There is already a tag. Append the text to its children.
                        if let Some(parent_tag) = stack.last_mut() {
                            parent_tag.add_text(text);
                        }
                    } else if let Some(tmp_str) = &mut str {
                        // We have a previous plain text string. Append to it.
                        tmp_str.push_str(text.as_str());
                    } else {
                        // Fresh plain text string.
                        str.replace(text);
                    }
                }

                Event::Start(xmlbytes) => {
                    if stack.is_empty() {
                        // No parent tag on the stack. Create the parent.
                        let mut toplevel = XmlTag::new(xml.decode(xml_tag.name())?);
                        copy_attr2(toplevel.attrmap_mut(), xml, xml_tag)?;
                        // Previous plain text strings are added.
                        if let Some(s) = &str {
                            toplevel.add_text(s.as_str());
                            str = None;
                        }
                        // Push to the stack.
                        stack.push(toplevel);
                    }

                    // Set the new tag.
                    let mut tag = XmlTag::new(xml.decode(xmlbytes.name())?);
                    copy_attr2(tag.attrmap_mut(), xml, &xmlbytes)?;
                    stack.push(tag);
                }

                Event::End(xmlbytes) => {
                    // End tag.
                    if xmlbytes.name() == end_tag {
                        if !stack.is_empty() {
                            assert_eq!(stack.len(), 1);
                            let tag = stack.pop().unwrap();
                            text.replace(tag);
                        }
                        break;
                    } else {
                        // Get the tag from the stack and add it to it's parent.
                        let tag = stack.pop().unwrap();
                        if let Some(parent) = stack.last_mut() {
                            parent.add_tag(tag);
                        } else {
                            unreachable!()
                        }
                    }
                }

                Event::Empty(xmlbytes) => {
                    if stack.is_empty() {
                        // No parent tag on the stack. Create the parent.
                        let mut toplevel = XmlTag::new(xml.decode(xml_tag.name())?);
                        copy_attr2(toplevel.attrmap_mut(), xml, xml_tag)?;
                        // Previous plain text strings are added.
                        if let Some(s) = &str {
                            toplevel.add_text(s.as_str());
                            str = None;
                        }
                        // Push to the stack.
                        stack.push(toplevel);
                    }

                    // Create the tag and append it immediately to the parent.
                    let mut emptytag = XmlTag::new(xml.decode(xmlbytes.name())?);
                    copy_attr2(emptytag.attrmap_mut(), xml, &xmlbytes)?;

                    if let Some(parent) = stack.last_mut() {
                        parent.add_tag(emptytag);
                    } else {
                        unreachable!()
                    }
                }

                Event::Eof => {
                    break;
                }

                _ => {
                    if cfg!(feature = "dump_unused") {
                        println!(" read_text_or_tag unused {:?}", evt);
                    }
                }
            }
        }
    }

    Ok((str, text))
}
