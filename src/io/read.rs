use std::fs::File;
use std::io::BufReader;
use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime};
use quick_xml::events::{BytesStart, Event};
use quick_xml::events::attributes::Attribute;
use zip::read::ZipFile;

use crate::{ColRange, RowRange, SCell, Sheet, ucell, Value, ValueFormat, ValueType, WorkBook};
use crate::attrmap::AttrMap;
use crate::error::OdsError;
use crate::format::{FormatPart, FormatPartType};
use crate::refs::{CellRef, parse_cellranges, parse_cellref};
use crate::style::{FontFaceDecl, HeaderFooter, PageLayout, Style, StyleFor, StyleMap, StyleOrigin, StyleUse};
use crate::text::{TextTag, TextVec};

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
fn read_content(book: &mut WorkBook,
                zip_file: &mut ZipFile) -> Result<(), OdsError> {

    // xml parser
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    xml.trim_text(true);

    let mut buf = Vec::new();

    let mut sheet = Sheet::new();

    // Separate counter for table-columns
    let mut tcol: ucell = 0;

    // Cell position
    let mut row: ucell = 0;
    let mut col: ucell = 0;

    // Rows can be repeated. In reality only empty ones ever are.
    let mut row_repeat: ucell = 1;
    // Row style.
    let mut row_style: Option<String> = None;

    let mut col_range_from = 0;
    let mut row_range_from = 0;

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_content {:?}", evt); }
        match evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:document-content"
                || xml_tag.name() == b"office:body"
                || xml_tag.name() == b"office:spreadsheet" => {
                // noop
            }

            Event::End(xml_tag)
            if xml_tag.name() == b"office:document-content"
                || xml_tag.name() == b"office:body"
                || xml_tag.name() == b"office:spreadsheet" => {
                // noop
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table" => {
                read_table_attr(&xml, xml_tag, &mut sheet)?;
            }

            Event::End(xml_tag)
            if xml_tag.name() == b"table:table" => {
                row = 0;
                col = 0;
                book.push_sheet(sheet);
                sheet = Sheet::new();
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-header-columns" => {
                col_range_from = tcol;
            }

            Event::End(xml_tag)
            if xml_tag.name() == b"table:table-header-columns" => {
                sheet.header_cols = Some(ColRange::new(col_range_from, tcol - 1));
            }

            Event::Empty(xml_tag)
            if xml_tag.name() == b"table:table-column" => {
                tcol = read_table_column(&mut xml, &xml_tag, tcol, &mut sheet)?;
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
                row_repeat = read_table_row_attr(&mut xml, xml_tag, &mut row_style)?;
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
                row_style = None;

                row += row_repeat;
                col = 0;
                row_repeat = 1;
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

            Event::Empty(xml_tag)
            if xml_tag.name() == b"table:table-cell" || xml_tag.name() == b"table:covered-table-cell" => {
                col = read_empty_table_cell(&mut sheet, &mut xml, xml_tag, row, col)?;
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"table:table-cell" || xml_tag.name() == b"table:covered-table-cell" => {
                col = read_table_cell(&mut sheet, &mut xml, xml_tag, row, col)?;
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

// Reads the table attributes.
fn read_table_attr(xml: &quick_xml::Reader<BufReader<&mut ZipFile>>,
                   xml_tag: BytesStart,
                   sheet: &mut Sheet) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:name" => {
                let v = attr.unescape_and_decode_value(xml)?;
                sheet.set_name(v);
            }
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(xml)?;
                sheet.set_style(v);
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
fn read_table_row_attr(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                       xml_tag: BytesStart,
                       row_style: &mut Option<String>) -> Result<ucell, OdsError>
{
    let mut row_repeat: ucell = 1;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:number-rows-repeated" => {
                let v = attr.unescaped_value()?;
                let v = xml.decode(v.as_ref())?;
                row_repeat = v.parse::<ucell>()?;
            }
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                *row_style = Some(v);
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

    Ok(row_repeat)
}

// Reads the table-column attributes. Creates as many copies as indicated.
fn read_table_column(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                     xml_tag: &BytesStart,
                     mut tcol: ucell,
                     sheet: &mut Sheet) -> Result<ucell, OdsError> {
    let mut style = None;
    let mut cell_style = None;
    let mut repeat: ucell = 1;

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style = Some(v);
            }
            attr if attr.key == b"table:number-columns-repeated" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                repeat = v.parse()?;
            }
            attr if attr.key == b"table:default-cell-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                cell_style = Some(v);
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
            sheet.set_column_style(tcol, style.clone());
        }
        if let Some(cell_style) = &cell_style {
            sheet.set_column_cell_style(tcol, cell_style.clone());
        }
        tcol += 1;
        repeat -= 1;
    }

    Ok(tcol)
}

// Reads the cell data.
fn read_table_cell(sheet: &mut Sheet,
                   xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                   xml_tag: BytesStart,
                   row: ucell,
                   mut col: ucell) -> Result<ucell, OdsError> {

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
    let mut cell_content_txt: Option<TextVec> = None;
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

            attr if attr.key == b"office:value-type" =>
                value_type = Some(decode_value_type(attr)?),
            attr if attr.key == b"calcext:value-type" =>
                {}

            attr if attr.key == b"office:date-value" =>
                cell_value = Some(attr.unescape_and_decode_value(&xml)?),
            attr if attr.key == b"office:time-value" =>
                cell_value = Some(attr.unescape_and_decode_value(&xml)?),
            attr if attr.key == b"office:value" =>
                cell_value = Some(attr.unescape_and_decode_value(&xml)?),
            attr if attr.key == b"office:boolean-value" =>
                cell_value = Some(attr.unescape_and_decode_value(&xml)?),

            attr if attr.key == b"office:currency" =>
                cell_currency = Some(attr.unescape_and_decode_value(&xml)?),

            attr if attr.key == b"table:formula" =>
                cell.formula = Some(attr.unescape_and_decode_value(&xml)?),
            attr if attr.key == b"table:style-name" =>
                cell.style = Some(attr.unescape_and_decode_value(&xml)?),

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
        if cfg!(feature = "dump_xml") { println!(" read_table_cell {:?}", evt); }
        match evt {
            Event::Start(xml_tag)
            if xml_tag.name() == b"text:p" => {
                read_textvec2(&mut cell_content_txt,
                              &mut cell_content,
                              b"text:p",
                              xml,
                )?;
            }

            Event::Empty(xml_tag)
            if xml_tag.name() == b"text:p" => {
                // noop
            }

            Event::End(xml_tag)
            if xml_tag.name() == tag_name => {
                cell.value = parse_value(value_type,
                                         cell_value,
                                         cell_content,
                                         cell_content_txt,
                                         cell_currency,
                                         row,
                                         col)?;

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
                if cfg!(feature = "dump_unused") { println!(" read_table_cell unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(col)
}

/// Reads a table-cell from an empty XML tag.
/// There seems to be no data associated, but it can have a style and a formula.
/// And first of all we need the repeat count for the correct placement.
fn read_empty_table_cell(sheet: &mut Sheet,
                         xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                         xml_tag: BytesStart,
                         row: ucell,
                         mut col: ucell) -> Result<ucell, OdsError> {
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
                cell.get_or_insert_with(SCell::new)
                    .set_style(attr.unescape_and_decode_value(&xml)?);
            }
            attr if attr.key == b"table:number-rows-spanned" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                let span = v.parse::<ucell>()?;

                cell.get_or_insert_with(SCell::new)
                    .set_row_span(span);
            }
            attr if attr.key == b"table:number-columns-spanned" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                let span = v.parse::<ucell>()?;

                cell.get_or_insert_with(SCell::new)
                    .set_col_span(span);
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
fn parse_value(value_type: Option<ValueType>,
               cell_value: Option<String>,
               cell_content: Option<String>,
               cell_content_txt: Option<TextVec>,
               cell_currency: Option<String>,
               row: ucell,
               col: ucell) -> Result<Value, OdsError> {
    if let Some(value_type) = value_type {
        match value_type {
            ValueType::Empty => {
                Ok(Value::Empty)
            }
            ValueType::Text => {
                if let Some(cell_content_txt) = cell_content_txt {
                    Ok(Value::TextM(cell_content_txt))
                } else if let Some(cell_content) = cell_content {
                    Ok(Value::Text(cell_content))
                } else {
                    Ok(Value::Text("".to_string()))
                }
            }
            ValueType::TextM => {
                unreachable!()
            }
            ValueType::Number => {
                if let Some(cell_value) = cell_value {
                    let f = cell_value.parse::<f64>()?;
                    Ok(Value::Number(f))
                } else {
                    Err(OdsError::Ods(format!("{} has type number, but no value!", CellRef::simple(row, col))))
                }
            }
            ValueType::DateTime => {
                if let Some(cell_value) = cell_value {
                    let dt =
                        if cell_value.len() == 10 {
                            NaiveDate::parse_from_str(cell_value.as_str(), "%Y-%m-%d")?.and_hms(0, 0, 0)
                        } else {
                            NaiveDateTime::parse_from_str(cell_value.as_str(), "%Y-%m-%dT%H:%M:%S%.f")?
                        };

                    Ok(Value::DateTime(dt))
                } else {
                    Err(OdsError::Ods(format!("{} has type datetime, but no value!", CellRef::simple(row, col))))
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
                    Err(OdsError::Ods(format!("{} has type time-duration, but no value!", CellRef::simple(row, col))))
                }
            }
            ValueType::Boolean => {
                if let Some(cell_value) = cell_value {
                    Ok(Value::Boolean(&cell_value == "true"))
                } else {
                    Err(OdsError::Ods(format!("{} has type boolean, but no value!", CellRef::simple(row, col))))
                }
            }
            ValueType::Currency => {
                if let Some(cell_value) = cell_value {
                    let f = cell_value.parse::<f64>()?;
                    if let Some(cell_currency) = cell_currency {
                        Ok(Value::Currency(cell_currency, f))
                    } else {
                        Err(OdsError::Ods(format!("{} has type currency, but no value!", CellRef::simple(row, col))))
                    }
                } else {
                    Err(OdsError::Ods(format!("{} has type currency, but no value!", CellRef::simple(row, col))))
                }
            }
            ValueType::Percentage => {
                if let Some(cell_value) = cell_value {
                    let f = cell_value.parse::<f64>()?;
                    Ok(Value::Percentage(f))
                } else {
                    Err(OdsError::Ods(format!("{} has type percentage, but no value!", CellRef::simple(row, col))))
                }
            }
        }
    } else {
        // could be an image or whatever
        Ok(Value::Empty)
    }
}

// String to ValueType
fn decode_value_type(attr: Attribute) -> Result<ValueType, OdsError> {
    match attr.unescaped_value()?.as_ref() {
        b"string" => Ok(ValueType::Text),
        b"float" => Ok(ValueType::Number),
        b"percentage" => Ok(ValueType::Percentage),
        b"date" => Ok(ValueType::DateTime),
        b"time" => Ok(ValueType::TimeDuration),
        b"boolean" => Ok(ValueType::Boolean),
        b"currency" => Ok(ValueType::Currency),
        other => Err(OdsError::Ods(format!("Unknown cell-type {:?}", other)))
    }
}

// reads a font-face
#[allow(clippy::single_match)]
fn read_fonts(book: &mut WorkBook,
              origin: StyleOrigin,
              xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut font: FontFaceDecl = FontFaceDecl::new_origin(origin);

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_fonts {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag)
            | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
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
                                    font.set_attr(k, v);
                                }
                            }
                        }

                        book.add_font(font);
                        font = FontFaceDecl::new_origin(StyleOrigin::Content);
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_fonts unused {:?}", evt); }
                    }
                }
            }

            Event::End(ref e) => {
                if e.name() == b"office:font-face-decls" {
                    break;
                }
            }

            Event::Eof => {
                break;
            }
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_fonts unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// reads the page-layout tag
fn read_page_layout(book: &mut WorkBook,
                    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                    xml_tag: &BytesStart) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut pl = PageLayout::default();
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

    let mut header_style = false;
    let mut footer_style = false;

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_page_layout {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag)
            | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:page-layout-properties" =>
                        copy_attr(&mut pl, xml, xml_tag)?,
                    b"style:header-style" =>
                        header_style = true,
                    b"style:footer-style" =>
                        footer_style = true,
                    b"style:header-footer-properties" => {
                        if header_style {
                            copy_attr(pl.header_attr_mut(), xml, xml_tag)?;
                        }
                        if footer_style {
                            copy_attr(pl.footer_attr_mut(), xml, xml_tag)?;
                        }
                    }
                    b"style:background-image" => {
                        // noop for now. sets the background transparent.
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_page_layout unused {:?}", evt); }
                    }
                }
            }
            Event::Text(_) => (),
            Event::End(ref end) => {
                match end.name() {
                    b"style:page-layout" =>
                        break,
                    b"style:header-style" =>
                        header_style = false,
                    b"style:footer-style" =>
                        footer_style = false,
                    b"style:header-footer-properties" =>
                        {}
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_page_layout unused {:?}", evt); }
                    }
                }
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_page_layout unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    book.add_pagelayout(pl);

    Ok(())
}

// read the master-styles tag
#[allow(clippy::single_match)]
fn read_master_styles(book: &mut WorkBook,
                      origin: StyleOrigin,
                      xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_master_styles {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag)
            | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:master-page" => {
                        read_master_page(book, origin, xml, xml_tag)?;
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_master_styles unused {:?}", evt); }
                    }
                }
            }
            Event::Text(_) => (),
            Event::End(ref e) => {
                if e.name() == b"office:master-styles" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_master_styles unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// read the master-page tag
fn read_master_page(book: &mut WorkBook,
                    _origin: StyleOrigin,
                    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                    xml_tag: &BytesStart) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut masterpage_name = "".to_string();
    let mut pagelayout_name = "".to_string();
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                masterpage_name = attr.unescape_and_decode_value(&xml)?;
            }
            attr if attr.key == b"style:page-layout-name" => {
                pagelayout_name = attr.unescape_and_decode_value(&xml)?;
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

    // may not exist? but should
    if book.pagelayout(&pagelayout_name).is_none() {
        let mut p = PageLayout::default();
        p.set_name(pagelayout_name.clone());
        book.add_pagelayout(p);
    }

    let pl = book.pagelayout_mut(&pagelayout_name).unwrap();
    pl.set_masterpage_name(masterpage_name);

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_master_page {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:header" => {
                        let hf = read_headerfooter(b"style:header", xml)?;
                        pl.set_header(hf);
                    }
                    b"style:header-left" => {
                        let hf = read_headerfooter(b"style:header-left", xml)?;
                        pl.set_header_left(hf);
                    }
                    b"style:footer" => {
                        let hf = read_headerfooter(b"style:footer", xml)?;
                        pl.set_footer(hf);
                    }
                    b"style:footer-left" => {
                        let hf = read_headerfooter(b"style:footer-left", xml)?;
                        pl.set_footer_left(hf);
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_master_page unused {:?}", evt); }
                    }
                }
            }

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
                if cfg!(feature = "dump_unused") { println!(" read_master_page unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// reads any header or footer tags
fn read_headerfooter(end_tag: &[u8],
                     xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<HeaderFooter, OdsError> {
    let mut buf = Vec::new();

    let mut hf = HeaderFooter::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_headerfooter {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag) |
            Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:region-left" => {
                        let cm = read_textvec(b"style:region-left", xml)?;
                        hf.set_left(cm);
                    }
                    b"style:region-center" => {
                        let cm = read_textvec(b"style:region-left", xml)?;
                        hf.set_center(cm);
                    }
                    b"style:region-right" => {
                        let cm = read_textvec(b"style:region-left", xml)?;
                        hf.set_right(cm);
                    }
                    b"text:p" => {
                        let cm = read_textvec(b"text:p", xml)?;
                        hf.set_content(cm);
                    }
                    // no other tags supported for now.
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_headerfooter unused {:?}", evt); }
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
                if cfg!(feature = "dump_unused") { println!(" read_headerfooter unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(hf)
}


// reads all the tags up to end_tag and creates a TextVec.
// if there are no tags the result is a plain String.
fn read_textvec2(vec: &mut Option<TextVec>,
                 str: &mut Option<String>,
                 end_tag: &[u8],
                 xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_textvec2 {:?}", evt); }
        match evt {
            Event::Text(ref txt) => {
                let t = txt.unescape_and_decode(xml)?;
                if let Some(vec) = vec {
                    vec.text(t);
                } else if let Some(s) = str {
                    let tmp = s.clone() + t.as_str();
                    str.replace(tmp);
                } else {
                    str.replace(t);
                }
            }

            Event::Start(ref xml_tag) => {
                if vec.is_none() {
                    let mut txtvec = TextVec::new();
                    if let Some(s) = str {
                        txtvec.text(s.as_str());
                        *str = None;
                    }
                    vec.replace(txtvec);
                }

                let mut c = TextTag::new(xml.decode(xml_tag.name())?);
                copy_attr(&mut c, xml, xml_tag)?;

                if let Some(vec) = vec {
                    vec.startc(c);
                }
            }

            Event::Empty(ref xml_tag) => {
                if vec.is_none() {
                    let mut txtvec = TextVec::new();
                    if let Some(s) = str {
                        txtvec.text(s.as_str());
                        *str = None;
                    }
                    vec.replace(txtvec);
                }

                let mut c = TextTag::new(xml.decode(xml_tag.name())?);
                copy_attr(&mut c, xml, xml_tag)?;

                if let Some(vec) = vec {
                    vec.emptyc(c);
                }
            }

            Event::End(ref xml_tag) => {
                if xml_tag.name() == end_tag {
                    break;
                } else {
                    if vec.is_none() {
                        let mut txtvec = TextVec::new();
                        if let Some(s) = str {
                            txtvec.text(s.as_str());
                            *str = None;
                        }
                        vec.replace(txtvec);
                    }

                    if let Some(vec) = vec {
                        vec.end(xml.decode(xml_tag.name())?);
                    }
                }
            }

            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_textvec2 unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// reads all the tags up to end_tag and creates a TextVec
fn read_textvec(end_tag: &[u8],
                xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<TextVec, OdsError> {
    let mut buf = Vec::new();

    let mut cv = TextVec::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_textvec {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag) => {
                let mut c = TextTag::new(xml.decode(xml_tag.name())?);
                copy_attr(&mut c, xml, xml_tag)?;
                cv.startc(c);
            }
            Event::Empty(ref xml_tag) => {
                let mut c = TextTag::new(xml.decode(xml_tag.name())?);
                copy_attr(&mut c, xml, xml_tag)?;
                cv.emptyc(c);
            }
            Event::Text(ref txt) => {
                cv.text(txt.unescape_and_decode(xml)?);
            }
            Event::End(ref xml_tag) => {
                if xml_tag.name() == end_tag {
                    break;
                } else {
                    cv.end(xml.decode(xml_tag.name())?);
                }
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_textvec unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(cv)
}

// reads the office-styles tag
fn read_styles_tag(book: &mut WorkBook,
                   origin: StyleOrigin,
                   xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = if let Event::Empty(_) = evt { true } else { false };
        if cfg!(feature = "dump_xml") { println!(" read_styles_tag {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag) |
            Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:style" => {
                        read_style_style(book, origin, StyleUse::Named, b"style:style", xml, xml_tag, empty_tag)?;
                    }
                    b"style:default-style" => {
                        read_style_style(book, origin, StyleUse::Default, b"style:default-style", xml, xml_tag, empty_tag)?;
                    }
                    b"number:boolean-style" |
                    b"number:date-style" |
                    b"number:time-style" |
                    b"number:number-style" |
                    b"number:currency-style" |
                    b"number:percentage-style" |
                    b"number:text-style" => {
                        read_value_format(book, origin, StyleUse::Named, xml, xml_tag)?;
                    }
                    // style:default-page-layout
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_styles_tag unused {:?}", evt); }
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
                if cfg!(feature = "dump_unused") { println!(" read_styles_tag unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// read the automatic-styles tag
fn read_auto_styles(book: &mut WorkBook,
                    origin: StyleOrigin,
                    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    loop {
        let evt = xml.read_event(&mut buf)?;
        let empty_tag = if let Event::Empty(_) = evt { true } else { false };
        if cfg!(feature = "dump_xml") { println!(" read_auto_styles {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag)
            | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"style:style" => {
                        read_style_style(book, origin, StyleUse::Automatic, b"style:style", xml, xml_tag, empty_tag)?;
                    }
                    b"number:boolean-style" |
                    b"number:date-style" |
                    b"number:time-style" |
                    b"number:number-style" |
                    b"number:currency-style" |
                    b"number:percentage-style" |
                    b"number:text-style" => {
                        read_value_format(book, origin, StyleUse::Automatic, xml, xml_tag)?;
                    }
                    // style:default-page-layout
                    b"style:page-layout" => {
                        read_page_layout(book, xml, xml_tag)?;
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_auto_styles unused {:?}", evt); }
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
                if cfg!(feature = "dump_unused") { println!(" read_auto_styles unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

// Reads any of the number:xxx tags
fn read_value_format(book: &mut WorkBook,
                     origin: StyleOrigin,
                     styleuse: StyleUse,
                     xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                     xml_tag: &BytesStart) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut value_style = ValueFormat::new_origin(origin, styleuse);
    // Styles with content information are stored before completion.
    let mut value_style_part = None;

    match xml_tag.name() {
        b"number:boolean-style" =>
            read_value_format_attr(ValueType::Boolean, &mut value_style, xml, xml_tag)?,
        b"number:date-style" =>
            read_value_format_attr(ValueType::DateTime, &mut value_style, xml, xml_tag)?,
        b"number:time-style" =>
            read_value_format_attr(ValueType::TimeDuration, &mut value_style, xml, xml_tag)?,
        b"number:number-style" =>
            read_value_format_attr(ValueType::Number, &mut value_style, xml, xml_tag)?,
        b"number:currency-style" =>
            read_value_format_attr(ValueType::Currency, &mut value_style, xml, xml_tag)?,
        b"number:percentage-style" =>
            read_value_format_attr(ValueType::Percentage, &mut value_style, xml, xml_tag)?,
        b"number:text-style" =>
            read_value_format_attr(ValueType::Text, &mut value_style, xml, xml_tag)?,
        _ => {
            if cfg!(feature = "dump_unused") {
                let n = xml.decode(xml_tag.name())?;
                println!(" read_value_format unused {}", n);
            }
        }
    }

    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_value_format {:?}", evt); }
        match evt {
            Event::Start(ref xml_tag)
            | Event::Empty(ref xml_tag) => {
                match xml_tag.name() {
                    b"number:boolean" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Boolean)?),
                    b"number:number" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Number)?),
                    b"number:scientific-number" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Scientific)?),
                    b"number:day" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Day)?),
                    b"number:month" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Month)?),
                    b"number:year" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Year)?),
                    b"number:era" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Era)?),
                    b"number:day-of-week" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::DayOfWeek)?),
                    b"number:week-of-year" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::WeekOfYear)?),
                    b"number:quarter" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Quarter)?),
                    b"number:hours" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Hours)?),
                    b"number:minutes" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Minutes)?),
                    b"number:seconds" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Seconds)?),
                    b"number:fraction" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Fraction)?),
                    b"number:am-pm" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::AmPm)?),
                    b"number:embedded-text" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::EmbeddedText)?),
                    b"number:text-content" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::TextContent)?),
                    b"style:text" =>
                        value_style.push_part(read_part(xml, xml_tag, FormatPartType::Day)?),
                    b"style:map" =>
                        value_style.push_stylemap(read_stylemap(xml, xml_tag)?),
                    b"number:currency-symbol" => {
                        value_style_part = Some(read_part(xml, xml_tag, FormatPartType::CurrencySymbol)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = value_style_part {
                                value_style.push_part(part);
                            }
                            value_style_part = None;
                        }
                    }
                    b"number:text" => {
                        value_style_part = Some(read_part(xml, xml_tag, FormatPartType::Text)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = value_style_part {
                                value_style.push_part(part);
                            }
                            value_style_part = None;
                        }
                    }
                    b"style:text-properties" =>
                        copy_attr(value_style.text_mut(), xml, xml_tag)?,
                    _ =>
                        if cfg!(feature = "dump_unused") { println!(" read_value_format unused {:?}", evt); }
                }
            }
            Event::Text(ref e) => {
                if let Some(part) = &mut value_style_part {
                    part.set_content(e.unescape_and_decode(&xml)?);
                }
            }
            Event::End(ref e) => {
                match e.name() {
                    b"number:boolean-style" |
                    b"number:date-style" |
                    b"number:time-style" |
                    b"number:number-style" |
                    b"number:currency-style" |
                    b"number:percentage-style" |
                    b"number:text-style" => {
                        book.add_format(value_style);
                        break;
                    }
                    b"number:currency-symbol" | b"number:text" => {
                        if let Some(part) = value_style_part {
                            value_style.push_part(part);
                        }
                        value_style_part = None;
                    }
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" read_value_format unused {:?}", evt); }
                    }
                }
            }
            Event::Eof => break,
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_value_format unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}

/// Copies all the attr from the tag.
fn read_value_format_attr(value_type: ValueType,
                          value_style: &mut ValueFormat,
                          xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                          xml_tag: &BytesStart) -> Result<(), OdsError> {
    value_style.set_value_type(value_type);

    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                value_style.set_name(v);
            }
            attr => {
                let k = xml.decode(&attr.key)?;
                let v = attr.unescape_and_decode_value(&xml)?;
                value_style.set_attr(k, v);
            }
        }
    }

    Ok(())
}

fn read_part(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
             xml_tag: &BytesStart,
             part_type: FormatPartType) -> Result<FormatPart, OdsError> {
    let mut part = FormatPart::new(part_type);
    copy_attr(&mut part, xml, xml_tag)?;
    Ok(part)
}

// style:style tag
#[allow(clippy::single_match)]
fn read_style_style(book: &mut WorkBook,
                    origin: StyleOrigin,
                    styleuse: StyleUse,
                    end_tag: &[u8],
                    xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                    xml_tag: &BytesStart,
                    empty_tag: bool) -> Result<(), OdsError> {
    let mut buf = Vec::new();
    let mut style: Style = Style::new_origin(origin, styleuse);

    read_style_attr(xml, xml_tag, &mut style)?;

    // In case of an empty xml-tag we are done here.
    if empty_tag {
        book.add_style(style);
    } else {
        loop {
            let evt = xml.read_event(&mut buf)?;
            if cfg!(feature = "dump_xml") { println!(" read_style_tag {:?}", evt); }
            match evt {
                Event::Start(ref xml_tag)
                | Event::Empty(ref xml_tag) => {
                    match xml_tag.name() {
                        b"style:table-properties" =>
                            copy_attr(style.table_mut(), xml, xml_tag)?,
                        b"style:table-column-properties" =>
                            copy_attr(style.col_mut(), xml, xml_tag)?,
                        b"style:table-row-properties" =>
                            copy_attr(style.row_mut(), xml, xml_tag)?,
                        b"style:table-cell-properties" =>
                            copy_attr(style.cell_mut(), xml, xml_tag)?,
                        b"style:text-properties" =>
                            copy_attr(style.text_mut(), xml, xml_tag)?,
                        b"style:paragraph-properties" =>
                            copy_attr(style.paragraph_mut(), xml, xml_tag)?,
                        b"style:map" =>
                            style.push_stylemap(read_stylemap(xml, xml_tag)?),
                        _ => {
                            if cfg!(feature = "dump_unused") { println!(" read_style_style unused {:?}", evt); }
                        }
                    }
                }
                Event::Text(_) => (),
                Event::End(ref e) => {
                    if e.name() == end_tag {
                        book.add_style(style);
                        break;
                    } else {
                        if cfg!(feature = "dump_unused") { println!(" read_style_style unused {:?}", evt); }
                    }
                }
                Event::Eof => break,
                _ => {
                    if cfg!(feature = "dump_unused") { println!(" read_style_style unused {:?}", evt); }
                }
            }
        }
    }

    Ok(())
}

fn read_stylemap(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                 xml_tag: &BytesStart) -> Result<StyleMap, OdsError> {
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
                    println!(" read_style_style unused {} {} {}", n, k, v);
                }
            }
        }
    }

    Ok(sm)
}

fn read_style_attr(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                   xml_tag: &BytesStart,
                   style: &mut Style) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        match attr? {
            attr if attr.key == b"style:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.set_name(v);
            }
            attr if attr.key == b"style:display-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.set_display_name(v);
            }
            attr if attr.key == b"style:family" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                match v.as_ref() {
                    "table" => style.set_family(StyleFor::Table),
                    "table-column" => style.set_family(StyleFor::TableColumn),
                    "table-row" => style.set_family(StyleFor::TableRow),
                    "table-cell" => style.set_family(StyleFor::TableCell),
                    _ => {
                        if cfg!(feature = "dump_unused") { println!(" style:family unused {} ", v); }
                    }
                }
            }
            attr if attr.key == b"style:parent-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.set_parent(v);
            }
            attr if attr.key == b"style:data-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.set_value_format(v);
            }
            attr => {
                if cfg!(feature = "dump_unused") {
                    let n = xml.decode(xml_tag.name())?;
                    let k = xml.decode(attr.key)?;
                    let v = attr.unescape_and_decode_value(xml)?;
                    println!(" read_style_attr unused {} {} {}", n, k, v);
                }
            }
        }
    }

    Ok(())
}

fn copy_attr(attrmap: &mut dyn AttrMap,
             xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
             xml_tag: &BytesStart) -> Result<(), OdsError> {
    for attr in xml_tag.attributes().with_checks(false) {
        if let Ok(attr) = attr {
            let k = xml.decode(&attr.key)?;
            let v = attr.unescape_and_decode_value(&xml)?;
            attrmap.set_attr(k, v);
        }
    }

    Ok(())
}

fn read_styles(book: &mut WorkBook,
               zip_file: &mut ZipFile) -> Result<(), OdsError> {
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    xml.trim_text(true);

    let mut buf = Vec::new();
    loop {
        let evt = xml.read_event(&mut buf)?;
        if cfg!(feature = "dump_xml") { println!(" read_styles {:?}", evt); }
        match evt {
            Event::Decl(_) => {}

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:document-styles" => {
                // noop
            }

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:font-face-decls" =>
                read_fonts(book, StyleOrigin::Styles, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:styles" =>
                read_styles_tag(book, StyleOrigin::Styles, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:automatic-styles" =>
                read_auto_styles(book, StyleOrigin::Styles, &mut xml)?,

            Event::Start(xml_tag)
            if xml_tag.name() == b"office:master-styles" =>
                read_master_styles(book, StyleOrigin::Styles, &mut xml)?,

            Event::Eof => {
                break;
            }
            _ => {
                if cfg!(feature = "dump_unused") { println!(" read_styles unused {:?}", evt); }
            }
        }

        buf.clear();
    }

    Ok(())
}



