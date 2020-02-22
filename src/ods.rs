use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::{File, rename};
use std::io;
use std::io::{BufReader, BufWriter, Read, Write};
use std::path::{Path, PathBuf};

use chrono::{Duration, NaiveDate, NaiveDateTime};
use log;
use quick_xml;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use zip;
use zip::read::ZipFile;
use zip::write::FileOptions;
use zip::ZipWriter;

use crate::{Family, Origin, Part, PartType, SCell, SColumn, Sheet, Style, Value, ValueStyle, ValueType, WorkBook};

#[derive(Debug)]
pub enum OdsError {
    Io(io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    ParseInt(std::num::ParseIntError),
    ParseFloat(std::num::ParseFloatError),
    Chrono(chrono::format::ParseError),
    Duration(time::OutOfRangeError),
    SystemTime(std::time::SystemTimeError),
}

impl From<time::OutOfRangeError> for OdsError {
    fn from(err: time::OutOfRangeError) -> OdsError {
        OdsError::Duration(err)
    }
}

impl From<io::Error> for OdsError {
    fn from(err: io::Error) -> OdsError {
        OdsError::Io(err)
    }
}

impl From<zip::result::ZipError> for OdsError {
    fn from(err: zip::result::ZipError) -> OdsError {
        OdsError::Zip(err)
    }
}

impl From<quick_xml::Error> for OdsError {
    fn from(err: quick_xml::Error) -> OdsError {
        OdsError::Xml(err)
    }
}

impl From<std::num::ParseIntError> for OdsError {
    fn from(err: std::num::ParseIntError) -> OdsError {
        OdsError::ParseInt(err)
    }
}

impl From<std::num::ParseFloatError> for OdsError {
    fn from(err: std::num::ParseFloatError) -> OdsError {
        OdsError::ParseFloat(err)
    }
}

impl From<chrono::format::ParseError> for OdsError {
    fn from(err: chrono::format::ParseError) -> OdsError {
        OdsError::Chrono(err)
    }
}

impl From<std::time::SystemTimeError> for OdsError {
    fn from(err: std::time::SystemTimeError) -> OdsError {
        OdsError::SystemTime(err)
    }
}

const DUMP_XML: bool = false;

// Reads an ODS-file.
pub fn read_ods(path: &Path) -> Result<WorkBook, OdsError> {
    let file = File::open(path)?;
    // ods is a zip-archive, we read content.xml
    let mut zip = zip::ZipArchive::new(file)?;
    let mut zip_file = zip.by_name("content.xml")?;

    let mut book = read_ods_content(&mut zip_file)?;

    book.file = Some(path.to_path_buf());

    Ok(book)
}

fn read_ods_content(zip_file: &mut ZipFile) -> Result<WorkBook, OdsError> {
    // xml parser
    let mut xml = quick_xml::Reader::from_reader(BufReader::new(zip_file));
    xml.trim_text(true);
    // xml.expand_empty_elements(true);

    let mut buf = Vec::new();

    let mut book = WorkBook::new();
    let mut sheet = Sheet::new();

    let mut tcol: usize = 0;
    let mut row: usize = 0;
    let mut col: usize = 0;
    // Empty rows are omitted and marked with a repeat-count.
    // Empty columns the same, but it's in an empty tag, we can handle there.
    let mut row_advance: usize = 1;
    // Columns can be repeated, not only empty ones.
    let mut col_repeat: usize = 1;

    let mut row_style: Option<String> = None;

    // Datatype in a cell
    let mut cell_type: String = String::from("");
    // Basic cell value here.
    let mut cell_value: Option<String> = None;
    // String content is held separately. It contains a formatted value of floats, dates etc
    let mut cell_string: Option<String> = None;
    // Currency + float-Value
    let mut cell_currency: Option<String> = None;
    // Formula if any.
    let mut cell_formula: Option<String> = None;
    let mut cell_style: Option<String> = None;

    loop {
        match xml.read_event(&mut buf)? {
            Event::Start(ref elem) => {
                if DUMP_XML { log::debug!("{:?}", elem); }

                match elem.name() {
                    b"table:table" => read_ods_table_tag(&mut xml, elem, &mut sheet)?,

                    b"table:table-row" => {
                        for a in elem.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"table:number-rows-repeated" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    row_advance = v.parse::<usize>()?;
                                }
                                Ok(ref attr) if attr.key == b"table:style-name" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    row_style = Some(v.to_string());
                                }
                                _ => {}
                            }
                        }
                    }
                    b"table:table-cell" => {
                        for a in elem.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"office:value-type" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_type = v.to_string();
                                }
                                Ok(ref attr) if attr.key == b"office:date-value" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_value = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"office:time-value" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_value = Some(v);
                                }
                                Ok(ref attr) if attr.key == b"office:value" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_value = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"office:boolean-value" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_value = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"office:currency" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_currency = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"table:formula" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_formula = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"table:style-name" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    cell_style = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"table:number-columns-repeated" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    col_repeat = v.parse::<usize>()?;
                                }
                                _ => {}
                            }
                        }
                    }
                    b"office:automatic-styles" => {
                        read_ods_styles(&mut book, &mut xml, b"office:automatic-styles")?;
                    }
                    _ => {}
                }
            }
            Event::End(ref elem) => {
                if DUMP_XML { log::debug!("{:?}", elem); }
                match elem.name() {
                    b"table:table" => {
                        row = 0;
                        col = 0;
                        book.push_sheet(sheet);
                        sheet = Sheet::new();
                    }
                    b"table:table-row" => {
                        if let Some(style) = row_style {
                            for r in row..row + row_advance {
                                sheet.set_row_style(r, style.clone());
                            }
                            row_style = None;
                        }

                        row += row_advance;
                        col = 0;
                        row_advance = 1;
                    }
                    b"table:table-cell" => {
                        let mut cell = SCell::new();

                        match cell_type.as_str() {
                            "string" => {
                                if let Some(cs) = cell_string {
                                    cell.value = Some(Value::from(cs));
                                }
                            }
                            "float" => {
                                if let Some(cs) = cell_value {
                                    let f = cs.parse::<f64>()?;
                                    cell.value = Some(Value::from(f));
                                }
                                cell_value = None;
                            }
                            "percentage" => {
                                if let Some(cs) = cell_value {
                                    let f = cs.parse::<f64>()?;
                                    cell.value = Some(Value::percentage(f));
                                }
                                cell_value = None;
                            }
                            "date" => {
                                if let Some(cs) = cell_value {
                                    let td =
                                        if cs.len() == 10 {
                                            NaiveDate::parse_from_str(cs.as_str(), "%Y-%m-%d")?.and_hms(0, 0, 0)
                                        } else {
                                            NaiveDateTime::parse_from_str(cs.as_str(), "%Y-%m-%dT%H:%M:%S%.f")?
                                        };
                                    cell.value = Some(Value::from(td));
                                }
                                cell_value = None;
                            }
                            "time" => {
                                if let Some(mut cs) = cell_value {
                                    let mut hour: u32 = 0;
                                    let mut have_hour = false;
                                    let mut min: u32 = 0;
                                    let mut have_min = false;
                                    let mut sec: u32 = 0;
                                    let mut have_sec = false;
                                    let mut nanos: u32 = 0;
                                    let mut nanos_digits: u8 = 0;

                                    for c in cs.drain(..) {
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

                                    cell.value = Some(Value::from(dur));
                                }
                                cell_value = None;
                            }
                            "boolean" => {
                                if let Some(cs) = cell_value {
                                    cell.value = Some(Value::from(&cs == "true"));
                                }
                                cell_value = None;
                            }
                            "currency" => {
                                if let Some(cs) = cell_value {
                                    let f = cs.parse::<f64>()?;
                                    cell.value = Some(Value::currency(&cell_currency.unwrap(), f));
                                }
                                cell_value = None;
                                cell_currency = None;
                            }
                            _ => {
                                log::warn!("Unknown cell-type {}", cell_type);
                            }
                        }

                        if let Some(formula) = cell_formula {
                            cell.formula = Some(formula);
                        }
                        cell_formula = None;
                        if let Some(style) = cell_style {
                            cell.style = Some(style);
                        }
                        cell_style = None;

                        while col_repeat > 1 {
                            sheet.add_cell(row, col, cell.clone());
                            col += 1;
                            col_repeat -= 1;
                        }
                        sheet.add_cell(row, col, cell);

                        cell_type = String::from("");
                        cell_string = None;
                        col += 1;
                    }
                    _ => {}
                }
            }
            Event::Empty(ref elem) => {
                if DUMP_XML { log::debug!("{:?}", elem); }
                match elem.name() {
                    b"table:table-column" => {
                        let mut column = SColumn::new();
                        let mut repeat: usize = 1;

                        for a in elem.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"table:style-name" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    column.style = Some(v.to_string());
                                }
                                Ok(ref attr) if attr.key == b"table:number-columns-repeated" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    repeat = v.parse()?;
                                }
                                Ok(ref attr) if attr.key == b"table:default-cell-style-name" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    column.def_cell_style = Some(v.to_string());
                                }
                                _ => {}
                            }
                        }

                        while repeat > 0 {
                            sheet.columns.insert(tcol, column.clone());
                            tcol += 1;
                            repeat -= 1;
                        }
                    }
                    b"table:table-cell" => {
                        if elem.attributes().count() == 0 {
                            col += 1;
                        }
                        for a in elem.attributes().with_checks(false) {
                            match a {
                                Ok(ref attr) if attr.key == b"table:number-columns-repeated" => {
                                    let v = attr.unescape_and_decode_value(&xml)?;
                                    col += v.parse::<usize>()?;
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Text(elem) => {
                if DUMP_XML { log::debug!("{:?}", elem); }
                let v = elem.unescape_and_decode(&xml)?;
                cell_string = Some(v);
            }
            Event::Eof => {
                if DUMP_XML { log::debug!("eof"); }
                break;
            }
            _ => {}
        }

        buf.clear();
    }

    Ok(book)
}

fn read_ods_table_tag(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                      xml_tag: &BytesStart,
                      sheet: &mut Sheet) -> Result<(), OdsError> {
    for a in xml_tag.attributes().with_checks(false) {
        match a {
            Ok(ref attr) if attr.key == b"table:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                sheet.set_name(v);
            }
            Ok(ref attr) if attr.key == b"table:style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                sheet.set_style(v);
            }
            _ => {}
        }
    }

    Ok(())
}

fn read_ods_styles(book: &mut WorkBook,
                   xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                   end_tag: &[u8]) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style: Style = Style::new(Origin::Content);
    let mut value_style = ValueStyle::new(Origin::Content);
    // Styles with content information are stored before completion.
    let mut value_style_part = None;

    loop {
        let evt = xml.read_event(&mut buf)?;
        match evt {
            Event::Start(ref xml_tag) | Event::Empty(ref xml_tag) => {
                if DUMP_XML { log::debug!(" style {:?}", xml_tag); }

                match xml_tag.name() {
                    b"style:style" => {
                        read_ods_style_tag(xml, xml_tag, &mut style)?;

                        // In case of an empty xml-tag we are done here.
                        if let Event::Empty(_) = evt {
                            book.add_style(style);
                            style = Style::new(Origin::Content);
                        }
                    }

                    b"style:table-properties" => read_ods_style_properties(xml, xml_tag, &mut style, &Style::set_table_prp)?,
                    b"style:table-column-properties" => read_ods_style_properties(xml, xml_tag, &mut style, &Style::set_table_col_prp)?,
                    b"style:table-row-properties" => read_ods_style_properties(xml, xml_tag, &mut style, &Style::set_table_row_prp)?,
                    b"style:table-cell-properties" => read_ods_style_properties(xml, xml_tag, &mut style, &Style::set_table_cell_prp)?,
                    b"style:text-properties" => read_ods_style_properties(xml, xml_tag, &mut style, &Style::set_text_prp)?,
                    b"style:paragraph-properties" => read_ods_style_properties(xml, xml_tag, &mut style, &Style::set_paragraph_prp)?,

                    b"number:boolean-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::Boolean, &mut value_style)?,
                    b"number:date-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::DateTime, &mut value_style)?,
                    b"number:time-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::TimeDuration, &mut value_style)?,
                    b"number:number-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::Number, &mut value_style)?,
                    b"number:currency-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::Currency, &mut value_style)?,
                    b"number:percentage-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::Percentage, &mut value_style)?,
                    b"number:text-style" => read_ods_value_style_tag(xml, xml_tag, ValueType::Text, &mut value_style)?,

                    b"number:boolean" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Boolean)?,
                    b"number:number" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Number)?,
                    b"number:scientific-number" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Scientific)?,
                    b"number:day" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Day)?,
                    b"number:month" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Month)?,
                    b"number:year" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Year)?,
                    b"number:era" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Era)?,
                    b"number:day-of-week" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::DayOfWeek)?,
                    b"number:week-of-year" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::WeekOfYear)?,
                    b"number:quarter" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Quarter)?,
                    b"number:hours" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Hours)?,
                    b"number:minutes" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Minutes)?,
                    b"number:seconds" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Seconds)?,
                    b"number:fraction" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Fraction)?,
                    b"number:am-pm" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::AmPm)?,
                    b"number:embedded-text" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::EmbeddedText)?,
                    b"number:text-content" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::TextContent)?,
                    b"style:text" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::Day)?,
                    b"style:map" => read_ods_value_style_part(xml, xml_tag, &mut value_style, PartType::StyleMap)?,
                    b"number:currency-symbol" => {
                        value_style_part = Some(read_ods_value_style_part2(xml, xml_tag, PartType::CurrencySymbol)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = value_style_part {
                                value_style.push_part(part);
                            }
                            value_style_part = None;
                        }
                    }
                    b"number:text" => {
                        value_style_part = Some(read_ods_value_style_part2(xml, xml_tag, PartType::Text)?);

                        // Empty-Tag. Finish here.
                        if let Event::Empty(_) = evt {
                            if let Some(part) = value_style_part {
                                value_style.push_part(part);
                            }
                            value_style_part = None;
                        }
                    }

                    _ => {}
                }
            }

            Event::Text(ref e) => {
                if DUMP_XML { log::debug!(" style {:?}", e); }
                if let Some(part) = &mut value_style_part {
                    part.content = Some(e.unescape_and_decode(&xml)?);
                }
            }

            Event::End(ref e) => {
                if DUMP_XML { log::debug!(" style {:?}", e); }

                if e.name() == end_tag {
                    break;
                }

                match e.name() {
                    b"style:style" => {
                        book.add_style(style);
                        style = Style::new(Origin::Content);
                    }
                    b"number:boolean-style" |
                    b"number:date-style" |
                    b"number:time-style" |
                    b"number:number-style" |
                    b"number:currency-style" |
                    b"number:percentage-style" |
                    b"number:text-style" => {
                        book.add_value_style(value_style);
                        value_style = ValueStyle::new(Origin::Content);
                    }
                    b"number:currency-symbol" | b"number:text" => {
                        if let Some(part) = value_style_part {
                            value_style.push_part(part);
                        }
                        value_style_part = None;
                    }

                    _ => {}
                }
            }
            Event::Eof => {
                if DUMP_XML { log::debug!("eof"); }
                break;
            }
            _ => {}
        }

        buf.clear();
    }

    Ok(())
}

fn read_ods_value_style_part(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                             xml_tag: &BytesStart,
                             value_style: &mut ValueStyle,
                             part_type: PartType) -> Result<(), OdsError> {
    value_style.push_part(read_ods_value_style_part2(xml, xml_tag, part_type)?);

    Ok(())
}

fn read_ods_value_style_part2(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                              xml_tag: &BytesStart,
                              part_type: PartType) -> Result<Part, OdsError> {
    let mut part = Part::new(part_type);

    for a in xml_tag.attributes().with_checks(false) {
        if let Ok(attr) = a {
            let k = xml.decode(&attr.key)?;
            let v = attr.unescape_and_decode_value(&xml)?;

            part.set_prp(k, v);
        }
    }

    Ok(part)
}

fn read_ods_value_style_tag(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                            xml_tag: &BytesStart,
                            value_type: ValueType,
                            value_style: &mut ValueStyle) -> Result<(), OdsError> {
    value_style.v_type = value_type;

    for a in xml_tag.attributes().with_checks(false) {
        match a {
            Ok(ref attr) if attr.key == b"style:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                value_style.set_name(v);
            }
            Ok(ref attr) => {
                let k = xml.decode(&attr.key)?;
                let v = attr.unescape_and_decode_value(&xml)?;
                value_style.set_prp(k, v);
            }
            _ => {}
        }
    }

    Ok(())
}

fn read_ods_style_properties(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                             xml_tag: &BytesStart,
                             style: &mut Style,
                             add_fn: &dyn Fn(&mut Style, &str, String)) -> Result<(), OdsError> {
    for a in xml_tag.attributes().with_checks(false) {
        if let Ok(attr) = a {
            let k = xml.decode(&attr.key)?;
            let v = attr.unescape_and_decode_value(&xml)?;
            add_fn(style, k, v);
        }
    }

    Ok(())
}

fn read_ods_style_tag(xml: &mut quick_xml::Reader<BufReader<&mut ZipFile>>,
                      xml_tag: &BytesStart,
                      style: &mut Style) -> Result<(), OdsError> {
    for a in xml_tag.attributes().with_checks(false) {
        match a {
            Ok(ref attr) if attr.key == b"style:name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.set_name(v);
            }
            Ok(ref attr) if attr.key == b"style:family" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                match v.as_ref() {
                    "table" => style.family = Family::Table,
                    "table-column" => style.family = Family::TableColumn,
                    "table-row" => style.family = Family::TableRow,
                    "table-cell" => style.family = Family::TableCell,
                    _ => {}
                }
            }
            Ok(ref attr) if attr.key == b"style:parent-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.parent = Some(v);
            }
            Ok(ref attr) if attr.key == b"style:data-style-name" => {
                let v = attr.unescape_and_decode_value(&xml)?;
                style.value_style = Some(v);
            }
            _ => { /* noop */ }
        }
    }

    Ok(())
}

fn xml_start(tag: &str) -> Event {
    let b = BytesStart::owned_name(tag.as_bytes());
    Event::Start(b)
}

fn xml_start_a<'a>(tag: &'a str, attr: Vec<(&'a str, String)>) -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr {
        b.push_attribute((a, v.as_ref()));
    }

    Event::Start(b)
}

fn xml_start_o<'a>(tag: &'a str, attr: Option<&'a HashMap<String, String>>) -> Event::<'a> {
    if let Some(attr) = attr {
        xml_start_m(tag, attr)
    } else {
        xml_start(tag)
    }
}

fn xml_start_m<'a>(tag: &'a str, attr: &'a HashMap<String, String>) -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr {
        b.push_attribute((a.as_str(), v.as_str()));
    }

    Event::Start(b)
}

fn xml_text(text: &str) -> Event {
    Event::Text(BytesText::from_plain_str(text))
}

fn xml_end(tag: &str) -> Event {
    let b = BytesEnd::borrowed(tag.as_bytes());
    Event::End(b)
}

fn xml_empty(tag: &str) -> Event {
    let b = BytesStart::owned_name(tag.as_bytes());
    Event::Empty(b)
}

fn xml_empty_a<'a>(tag: &'a str, attr: Vec<(&'a str, String)>) -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr {
        b.push_attribute((a, v.as_ref()));
    }

    Event::Empty(b)
}

fn xml_empty_o<'a>(tag: &'a str, attr: Option<&'a HashMap<String, String>>) -> Event::<'a> {
    if let Some(attr) = attr {
        xml_empty_m(tag, attr)
    } else {
        xml_empty(tag)
    }
}

fn xml_empty_m<'a>(tag: &'a str, attr: &'a HashMap<String, String>) -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr.iter() {
        b.push_attribute((a.as_str(), v.as_str()));
    }

    Event::Empty(b)
}

/// Writes the ODS file.
pub fn write_ods<P: AsRef<Path>>(book: &WorkBook, ods_path: P) -> Result<(), OdsError> {
    let orig_bak = if let Some(ods_orig) = &book.file {
        let mut orig_bak = ods_orig.clone();
        orig_bak.set_extension("bak");
        rename(&ods_orig, &orig_bak)?;
        Some(orig_bak)
    } else {
        None
    };

    let zip_file = File::create(ods_path)?;
    let mut zip_writer = zip::ZipWriter::new(BufWriter::new(zip_file));

    let mut file_set = HashSet::<String>::new();

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

    Ok(())
}

fn copy_workbook(ods_orig_name: &PathBuf, file_set: &mut HashSet<String>, zip_writer: &mut ZipWriter<BufWriter<File>>) -> Result<(), OdsError> {
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

fn write_mimetype(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), io::Error> {
    if !file_set.contains("mimetype") {
        file_set.insert(String::from("mimetype"));

        zip_out.start_file("mimetype", FileOptions::default().compression_method(zip::CompressionMethod::Stored))?;

        let mime = "application/vnd.oasis.opendocument.spreadsheet";
        zip_out.write_all(mime.as_bytes())?;
    }

    Ok(())
}

fn write_manifest(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("META-INF/manifest.xml") {
        file_set.insert(String::from("META-INF/manifest.xml"));

        zip_out.add_directory("META-INF", FileOptions::default())?;
        zip_out.start_file("META-INF/manifest.xml", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new_with_indent(zip_out, 32, 1);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;

        xml_out.write_event(xml_start_a("manifest:manifest", vec![
            ("xmlns:manifest", String::from("urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")),
            ("manifest:version", String::from("1.2")),
        ]))?;

        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("/")),
            ("manifest:version", String::from("1.2")),
            ("manifest:media-type", String::from("application/vnd.oasis.opendocument.spreadsheet")),
        ]))?;
//        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
//            ("manifest:full-path", String::from("Configurations2/")),
//            ("manifest:media-type", String::from("application/vnd.sun.xml.ui.configuration")),
//        ]))?;
        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("manifest.rdf")),
            ("manifest:media-type", String::from("application/rdf+xml")),
        ]))?;
        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("styles.xml")),
            ("manifest:media-type", String::from("text/xml")),
        ]))?;
        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("meta.xml")),
            ("manifest:media-type", String::from("text/xml")),
        ]))?;
        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
            ("manifest:full-path", String::from("content.xml")),
            ("manifest:media-type", String::from("text/xml")),
        ]))?;
//        xml_out.write_event(xml_empty_a("manifest:file-entry", vec![
//            ("manifest:full-path", String::from("settings.xml")),
//            ("manifest:media-type", String::from("text/xml")),
//        ]))?;
        xml_out.write_event(xml_end("manifest:manifest"))?;
    }

    Ok(())
}

fn write_manifest_rdf(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("manifest.rdf") {
        file_set.insert(String::from("manifest.rdf"));

        zip_out.start_file("manifest.rdf", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new(zip_out);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
        xml_out.write(b"\n")?;

        xml_out.write_event(xml_start_a("rdf:RDF", vec![
            ("xmlns:rdf", String::from("http://www.w3.org/1999/02/22-rdf-syntax-ns#")),
        ]))?;

        xml_out.write_event(xml_start_a("rdf:Description", vec![
            ("rdf:about", String::from("content.xml")),
        ]))?;
        xml_out.write_event(xml_empty_a("rdf:type", vec![
            ("rdf:resource", String::from("http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile")),
        ]))?;
        xml_out.write_event(xml_end("rdf:Description"))?;

        xml_out.write_event(xml_start_a("rdf:Description", vec![
            ("rdf:about", String::from("")),
        ]))?;
        xml_out.write_event(xml_empty_a("ns0:hasPart", vec![
            ("xmlns:ns0", String::from("http://docs.oasis-open.org/ns/office/1.2/meta/pkg#")),
            ("rdf:resource", String::from("content.xml")),
        ]))?;
        xml_out.write_event(xml_end("rdf:Description"))?;

        xml_out.write_event(xml_start_a("rdf:Description", vec![
            ("rdf:about", String::from("")),
        ]))?;
        xml_out.write_event(xml_empty_a("rdf:type", vec![
            ("rdf:resource", String::from("http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document")),
        ]))?;
        xml_out.write_event(xml_end("rdf:Description"))?;

        xml_out.write_event(xml_end("rdf:RDF"))?;
    }

    Ok(())
}

fn write_meta(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("meta.xml") {
        file_set.insert(String::from("meta.xml"));

        zip_out.start_file("meta.xml", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new(zip_out);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
        xml_out.write(b"\n")?;

        xml_out.write_event(xml_start_a("office:document-meta", vec![
            ("xmlns:meta", String::from("urn:oasis:names:tc:opendocument:xmlns:meta:1.0")),
            ("xmlns:office", String::from("urn:oasis:names:tc:opendocument:xmlns:office:1.0")),
            ("office:version", String::from("1.2")),
        ]))?;

        xml_out.write_event(xml_start("office:meta"))?;

        xml_out.write_event(xml_start("meta:generator"))?;
        xml_out.write_event(xml_text("spreadsheet-ods 0.1.0"))?;
        xml_out.write_event(xml_end("meta:generator"))?;

        xml_out.write_event(xml_start("meta:creation-date"))?;
        let s = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH)?;
        let d = NaiveDateTime::from_timestamp(s.as_secs() as i64, 0);
        xml_out.write_event(xml_text(&d.format("%Y-%m-%dT%H:%M:%S%.f").to_string()))?;
        xml_out.write_event(xml_end("meta:creation-date"))?;

        xml_out.write_event(xml_start("meta:editing-duration"))?;
        xml_out.write_event(xml_text("P0D"))?;
        xml_out.write_event(xml_end("meta:editing-duration"))?;

        xml_out.write_event(xml_start("meta:editing-cycles"))?;
        xml_out.write_event(xml_text("1"))?;
        xml_out.write_event(xml_end("meta:editing-cycles"))?;

        xml_out.write_event(xml_start("meta:initial-creator"))?;
        xml_out.write_event(xml_text(&username::get_user_name().unwrap()))?;
        xml_out.write_event(xml_end("meta:initial-creator"))?;

        xml_out.write_event(xml_end("office:meta"))?;

        xml_out.write_event(xml_end("office:document-meta"))?;
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
//        xml_out.write_event(xml_end("office:settings"))?;
//
//        xml_out.write_event(xml_end("office:document-settings"))?;
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

fn write_ods_styles(zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    if !file_set.contains("styles.xml") {
        file_set.insert(String::from("styles.xml"));

        zip_out.start_file("styles.xml", FileOptions::default())?;

        let mut xml_out = quick_xml::Writer::new(zip_out);

        xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
        xml_out.write(b"\n")?;

        xml_out.write_event(xml_start_a("office:document-styles", vec![
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

        // TODO

        xml_out.write_event(xml_end("office:document-styles"))?;
    }

    Ok(())
}

fn write_ods_content(book: &WorkBook, zip_out: &mut ZipWriter<BufWriter<File>>, file_set: &mut HashSet<String>) -> Result<(), OdsError> {
    file_set.insert(String::from("content.xml"));

    zip_out.start_file("content.xml", FileOptions::default())?;

    let mut xml_out = quick_xml::Writer::new(zip_out);

    xml_out.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
    xml_out.write(b"\n")?;
    xml_out.write_event(xml_start_a("office:document-content", vec![
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
    xml_out.write_event(xml_empty("office:scripts"))?;
    xml_out.write_event(xml_empty("office:font-face-decls"))?;

    xml_out.write_event(xml_start("office:automatic-styles"))?;
    write_styles(&book.styles, Origin::Content, &mut xml_out)?;
    write_value_styles(&book.value_styles, Origin::Content, &mut xml_out)?;
    xml_out.write_event(xml_end("office:automatic-styles"))?;

    xml_out.write_event(xml_start("office:body"))?;
    xml_out.write_event(xml_start("office:spreadsheet"))?;

    for sheet in &book.sheets {
        write_sheet(&book, &sheet, &mut xml_out)?;
    }

    xml_out.write_event(xml_end("office:spreadsheet"))?;
    xml_out.write_event(xml_end("office:body"))?;
    xml_out.write_event(xml_end("office:document-content"))?;

    Ok(())
}

fn write_sheet(book: &WorkBook, sheet: &Sheet, xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    attr.push(("table-name", sheet.name.to_string()));
    if let Some(style) = &sheet.style {
        attr.push(("table:style-name", style.to_string()));
    }
    xml_out.write_event(xml_start_a("table:table", attr))?;

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
            xml_out.write_event(xml_end("table:table-row"))?;
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
        xml_out.write_event(xml_end("table:table-row"))?;
    }

    xml_out.write_event(xml_end("table:table"))?;

    Ok(())
}

fn write_empty_cells(forward_dc: i32,
                     xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    attr.push(("table:number-columns-repeated", (forward_dc - 1).to_string()));
    xml_out.write_event(xml_empty_a("table:table-cell", attr))?;

    Ok(())
}

fn write_start_current_row(sheet: &Sheet,
                           cur_row: usize,
                           backward_dc: i32,
                           xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
    let mut attr = Vec::new();
    if let Some(row) = sheet.rows.get(&cur_row) {
        if let Some(style) = &row.style {
            attr.push(("table:style-name", style.to_string()));
        }
    }

    xml_out.write_event(xml_start_a("table:table-row", attr))?;

    // Might not be the first column in this row.
    if backward_dc > 0 {
        xml_out.write_event(xml_empty_a("table:table-cell", vec![
            ("table:number-columns-repeated", backward_dc.to_string()),
        ]))?;
    }

    Ok(())
}

fn write_empty_rows_before(first_cell: bool,
                           backward_dr: i32,
                           max_cell: &(usize, usize),
                           xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
    // Empty rows in between are 1 less than the delta, except at the very start.
    let mut attr = Vec::new();

    let empty_count = if first_cell { backward_dr } else { backward_dr - 1 };
    if empty_count > 1 {
        attr.push(("table:number-rows-repeated", empty_count.to_string()));
    }

    xml_out.write_event(xml_start_a("table:table-row", attr))?;

    // We fill the empty spaces completely up to max columns.
    xml_out.write_event(xml_empty_a("table:table-cell", vec![
        ("table:number-columns-repeated", max_cell.1.to_string()),
    ]))?;

    xml_out.write_event(xml_end("table:table-row"))?;

    Ok(())
}

fn write_table_columns(sheet: &Sheet,
                       max_cell: &(usize, usize),
                       xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
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
        xml_out.write_event(xml_empty_a("table:table-column", attr))?;
    }

    Ok(())
}

fn write_cell(book: &WorkBook,
              cell: &SCell,
              xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
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
        xml_out.write_event(xml_start_a("table:table-cell", attr))?;
        xml_out.write_event(xml_start("text:p"))?;
        xml_out.write_event(xml_text(&content))?;
        xml_out.write_event(xml_end("text:p"))?;
        xml_out.write_event(xml_end("table:table-cell"))?;
    } else {
        xml_out.write_event(xml_empty_a("table:table-cell", attr))?;
    }

    Ok(())
}

fn write_styles(styles: &BTreeMap<String, Style>, origin: Origin, xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
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
        xml_out.write_event(xml_start_a("style:style", attr))?;

        if let Some(table_cell_prp) = &style.table_cell_prp {
            xml_out.write_event(xml_empty_m("style:table-cell-properties", &table_cell_prp))?;
        }
        if let Some(table_col_prp) = &style.table_col_prp {
            xml_out.write_event(xml_empty_m("style:table-column-properties", &table_col_prp))?;
        }
        if let Some(table_row_prp) = &style.table_row_prp {
            xml_out.write_event(xml_empty_m("style:table-row-properties", &table_row_prp))?;
        }
        if let Some(table_prp) = &style.table_prp {
            xml_out.write_event(xml_empty_m("style:table-properties", &table_prp))?;
        }
        if let Some(paragraph_prp) = &style.paragraph_prp {
            xml_out.write_event(xml_empty_m("style:paragraph-properties", &paragraph_prp))?;
        }
        if let Some(text_prp) = &style.text_prp {
            xml_out.write_event(xml_empty_m("style:text-properties", &text_prp))?;
        }

        xml_out.write_event(xml_end("style:style"))?;
    }

    Ok(())
}

fn write_value_styles(styles: &BTreeMap<String, ValueStyle>, origin: Origin, xml_out: &mut quick_xml::Writer<&mut ZipWriter<BufWriter<File>>>) -> Result<(), OdsError> {
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

        let mut attr: HashMap<String, String> = HashMap::new();
        attr.insert(String::from("style:name"), style.name.to_string());
        if let Some(m) = &style.prp {
            for (k, v) in m {
                attr.insert(k.to_string(), v.to_string());
            }
        }
        xml_out.write_event(xml_start_m(tag, &attr))?;

        if let Some(parts) = style.parts() {
            for part in parts {
                let part_tag = match part.p_type {
                    PartType::Boolean => "number:boolean",
                    PartType::Number => "number:number",
                    PartType::Scientific => "number:scientific-number",
                    PartType::CurrencySymbol => "number:currency-symbol",
                    PartType::Day => "number:day",
                    PartType::Month => "number:month",
                    PartType::Year => "number:year",
                    PartType::Era => "number:era",
                    PartType::DayOfWeek => "number:day-of-week",
                    PartType::WeekOfYear => "number:week-of-year",
                    PartType::Quarter => "number:quarter",
                    PartType::Hours => "number:hours",
                    PartType::Minutes => "number:minutes",
                    PartType::Seconds => "number:seconds",
                    PartType::Fraction => "number:fraction",
                    PartType::AmPm => "number:am-pm",
                    PartType::EmbeddedText => "number:embedded-text",
                    PartType::Text => "number:text",
                    PartType::TextContent => "number:text-content",
                    PartType::StyleText => "style:text",
                    PartType::StyleMap => "style:map",
                };

                if part.p_type == PartType::Text || part.p_type == PartType::CurrencySymbol {
                    xml_out.write_event(xml_start_o(part_tag, part.prp.as_ref()))?;
                    if let Some(content) = &part.content {
                        xml_out.write_event(xml_text(content))?;
                    }
                    xml_out.write_event(xml_end(part_tag))?;
                } else {
                    xml_out.write_event(xml_empty_o(part_tag, part.prp.as_ref()))?;
                }
            }
        }
        xml_out.write_event(xml_end(tag))?;
    }

    Ok(())
}
