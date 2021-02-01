use std::fs::File;
use std::io::BufReader;
use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime};
use quick_xml::events::{BytesStart, Event};
use zip::read::ZipFile;

use crate::attrmap2::{AttrMap2, AttrMap2Trait};
use crate::error::OdsError;
use crate::format::{FormatPart, FormatPartType};
use crate::refs::{parse_cellranges, parse_cellref, CellRef};
use crate::style::stylemap::StyleMap;
use crate::style::tabstop::TabStop;
use crate::style::{
    ColStyle, FontFaceDecl, GraphicStyle, HeaderFooter, MasterPage, PageStyle, ParagraphStyle,
    RowStyle, StyleOrigin, StyleUse, TableStyle, TextStyle,
};
use crate::text::TextTag;
use crate::xmltree::{XmlContent, XmlTag};
use crate::{
    ucell, CellStyle, ColRange, RowRange, SCell, Sheet, Value, ValueFormat, ValueType, Visibility,
    WorkBook,
};

/// Reads an ODS-file.
pub fn read_ods<P: AsRef<Path>>(path: P) -> Result<WorkBook, OdsError> {
    let file = File::open(path.as_ref())?;
    // ods is a zip-archive, we read content.xml
    let mut zip = zip::ZipArchive::new(file)?;

    let mut book = WorkBook::new();
    book.file = Some(path.as_ref().to_path_buf());

    read_content(&mut book, &mut zip.by_name("content.xml")?)?;
    read_styles(&mut book, &mut zip.by_name("styles.xml")?)?;

    Ok(book)
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
                xml_tag.name() == b"table:content-validations" ||
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
                xml_tag.name() == b"table:content-validations" ||
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
                // There is often a strange repeat count for the last
                // row of the table that is in the millions.
                // That hits the break quite thoroughly, for now I ignore
                // this. Removes the row style for empty rows, I can live
                // with that for now.
                //
                // if let Some(style) = row_style {
                //     for r in row..row + row_repeat {
                //         sheet.set_row_style(r, style.clone());
                //     }
                // }
                if let Some(rowstyle) = rowstyle {
                    sheet.set_rowstyle(row, &rowstyle.into());
                }
                rowstyle = None;
                if let Some(row_cellstyle) = row_cellstyle {
                    sheet.set_row_cellstyle(row, &row_cellstyle.into());
                }
                row_cellstyle = None;
                sheet.set_row_visible(row, row_visible);
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
    let mut cell_content_txt: Option<TextTag> = None;
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
                cell_content = str;
                cell_content_txt = txt;
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
    cell_content_txt: Option<TextTag>,
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
                    Ok(Value::TextXml(Box::new(cell_content_txt)))
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
                    let mut hour: u32 = 0;
                    let mut have_hour = false;
                    let mut min: u32 = 0;
                    let mut have_min = false;
                    let mut sec: u32 = 0;
                    let mut have_sec = false;
                    let mut nanos: u32 = 0;
                    let mut nanos_digits: u8 = 0;

                    for c in cell_value.drain(..) {
                        match c {
                            'P' | 'T' => {}
                            '0'..='9' => {
                                if !have_hour {
                                    hour = hour * 10 + (c as u32 - '0' as u32);
                                } else if !have_min {
                                    min = min * 10 + (c as u32 - '0' as u32);
                                } else if !have_sec {
                                    sec = sec * 10 + (c as u32 - '0' as u32);
                                } else {
                                    nanos = nanos * 10 + (c as u32 - '0' as u32);
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

                    let secs: u64 = hour as u64 * 3600 + min as u64 * 60 + sec as u64;
                    let dur = Duration::from_std(std::time::Duration::new(secs, nanos))?;

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

// read the master-styles tag
fn read_master_styles(
    book: &mut WorkBook,
    origin: StyleOrigin,
    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
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
                    masterpage.set_header(read_headerfooter(b"style:header", xml)?);
                }
                b"style:header-first" => {
                    masterpage.set_header_first(read_headerfooter(b"style:header-first", xml)?);
                }
                b"style:header-left" => {
                    masterpage.set_header_left(read_headerfooter(b"style:header-left", xml)?);
                }
                b"style:footer" => {
                    masterpage.set_footer(read_headerfooter(b"style:footer", xml)?);
                }
                b"style:footer-first" => {
                    masterpage.set_footer_first(read_headerfooter(b"style:footer-first", xml)?);
                }
                b"style:footer-left" => {
                    masterpage.set_footer_left(read_headerfooter(b"style:footer-left", xml)?);
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
) -> Result<HeaderFooter, OdsError> {
    let mut buf = Vec::new();

    let mut hf = HeaderFooter::new();

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
#[allow(clippy::collapsible_if)]
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
#[allow(clippy::collapsible_if)]
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
#[allow(clippy::collapsible_if)]
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
#[allow(clippy::collapsible_if)]
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
#[allow(clippy::collapsible_if)]
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
#[allow(clippy::collapsible_if)]
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
#[allow(clippy::collapsible_if)]
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
                sm.set_condition(v);
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
                match v.as_ref() {
                    "table" | "table-column" | "table-row" | "table-cell" | "graphic"
                    | "paragraph" | "text" => return Ok(v.as_str().to_string()),
                    _ => {
                        return Err(OdsError::Ods(format!("style:family unknown {} ", v)));
                    }
                }
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
    for attr in xml_tag.attributes().with_checks(false) {
        if let Ok(attr) = attr {
            let k = xml.decode(&attr.key)?;
            let v = attr.unescape_and_decode_value(&xml)?;
            attrmap.set_attr(k, v);
        }
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
