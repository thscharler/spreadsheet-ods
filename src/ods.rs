use std::collections::{BTreeMap, HashMap};
use std::fs::File;
use std::io;
use std::io::{BufReader, BufWriter, Read, Write};
use std::path::Path;
use std::rc::Rc;

use chrono::{Duration, NaiveDate, NaiveDateTime};
use log;
use quick_xml::{Reader, Writer};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use tempfile::TempDir;
use walkdir::WalkDir;
use zip;
use zip::read::ZipFile;
use zip::write::FileOptions;

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
            Event::Start(ref e) => {
                if DUMP_XML { log::debug!("{:?}", e); }

                match e.name() {
                    b"table:table" => {
                        for a in e.attributes().with_checks(false) {
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
                    }
                    b"table:table-row" => {
                        for a in e.attributes().with_checks(false) {
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
                        for a in e.attributes().with_checks(false) {
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
                                    cell_value = Some(v.to_string());
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
            Event::End(ref e) => {
                if DUMP_XML { log::debug!("{:?}", e); }
                match e.name() {
                    b"table:table" => {
                        row = 0;
                        col = 0;
                        book.push_sheet(sheet);
                        sheet = Sheet::new();
                    }
                    b"table:table-row" => {
                        if let Some(style) = row_style {
                            let style_rc = Rc::new(style);
                            for r in row..row + row_advance {
                                sheet.set_row_style_rc(r, Rc::clone(&style_rc));
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
                                    let td;
                                    if cs.len() == 10 {
                                        td = NaiveDate::parse_from_str(cs.as_str(), "%Y-%m-%d")?.and_hms(0, 0, 0);
                                    } else {
                                        td = NaiveDateTime::parse_from_str(cs.as_str(), "%Y-%m-%dT%H:%M:%S%.f")?;
                                    }
                                    cell.value = Some(Value::from(td));
                                }
                                cell_value = None;
                            }
                            "time" => {
                                if let Some(mut cs) = cell_value {
                                    let mut h: u32 = 0;
                                    let mut have_h = false;
                                    let mut m: u32 = 0;
                                    let mut have_m = false;
                                    let mut s: u32 = 0;
                                    let mut have_s = false;
                                    let mut n: u32 = 0;
                                    let mut cn: u8 = 0;

                                    for c in cs.drain(..) {
                                        match c {
                                            'P' | 'T' => {}
                                            '0'..='9' => {
                                                if !have_h {
                                                    h = h * 10 + (c as u32 - '0' as u32);
                                                } else if !have_m {
                                                    m = m * 10 + (c as u32 - '0' as u32);
                                                } else if !have_s {
                                                    s = s * 10 + (c as u32 - '0' as u32);
                                                } else {
                                                    n = n * 10 + (c as u32 - '0' as u32);
                                                    cn += 1;
                                                }
                                            }
                                            'H' => have_h = true,
                                            'M' => have_m = true,
                                            '.' => have_s = true,
                                            'S' => {}
                                            _ => {}
                                        }
                                    }
                                    // unseen nano digits
                                    while cn < 9 {
                                        n = n * 10;
                                        cn += 1;
                                    }

                                    let secs: u64 = h as u64 * 3600 + m as u64 * 60 + s as u64;
                                    let dur = Duration::from_std(std::time::Duration::new(secs, n))?;

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
            Event::Empty(ref e) => {
                if DUMP_XML { log::debug!("{:?}", e); }
                match e.name() {
                    b"table:table-column" => {
                        let mut column = SColumn::new();
                        let mut repeat: usize = 1;

                        for a in e.attributes().with_checks(false) {
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
                        if e.attributes().count() == 0 {
                            col += 1;
                        }
                        for a in e.attributes().with_checks(false) {
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
            Event::Text(e) => {
                if DUMP_XML { log::debug!("{:?}", e); }
                let v = e.unescape_and_decode(&xml)?;
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

fn read_ods_styles(book: &mut WorkBook,
                   xml: &mut Reader<BufReader<&mut ZipFile>>,
                   end_tag: &[u8]) -> Result<(), OdsError> {
    let mut buf = Vec::new();

    let mut style: Style = Style::new(Origin::Content);
    let mut value_style = ValueStyle::new(Origin::Content);

    // String content is held separately. It contains a formatted value of floats, dates etc
    let mut cell_string: Option<String> = None;
    let mut number_attr: Option<HashMap<String, String>> = None;

    loop {
        let evt = xml.read_event(&mut buf)?;
        match evt {
            Event::Start(ref e) | Event::Empty(ref e) => {
                if DUMP_XML { log::debug!(" style {:?}", e); }

                match e.name() {
                    b"style:style" => {
                        for a in e.attributes().with_checks(false) {
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

                        // In case of an empty xml-tag we are done here.
                        if let Event::Empty(_) = evt {
                            book.add_style(style);
                            style = Style::new(Origin::Content);
                        }
                    }
                    b"style:table-properties" => {
                        for a in e.attributes().with_checks(false) {
                            if let Ok(attr) = a {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                style.set_table_prp(k, v);
                            }
                        }
                    }
                    b"style:table-column-properties" => {
                        for a in e.attributes().with_checks(false) {
                            if let Ok(attr) = a {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                style.set_table_col_prp(k, v);
                            }
                        }
                    }
                    b"style:table-row-properties" => {
                        for a in e.attributes().with_checks(false) {
                            if let Ok(attr) = a {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                style.set_table_row_prp(k, v);
                            }
                        }
                    }
                    b"style:table-cell-properties" => {
                        for a in e.attributes().with_checks(false) {
                            if let Ok(attr) = a {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                style.set_table_cell_prp(k, v);
                            }
                        }
                    }
                    b"style:text-properties" => {
                        for a in e.attributes().with_checks(false) {
                            if let Ok(attr) = a {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                style.set_text_prp(k, v);
                            }
                        }
                    }
                    b"style:paragraph-properties" => {
                        for a in e.attributes().with_checks(false) {
                            if let Ok(attr) = a {
                                let k = xml.decode(&attr.key)?;
                                let v = attr.unescape_and_decode_value(&xml)?;
                                style.set_paragraph_prp(k, v);
                            }
                        }
                    }

                    b"number:boolean-style" |
                    b"number:date-style" |
                    b"number:time-style" |
                    b"number:number-style" |
                    b"number:currency-style" |
                    b"number:percentage-style" |
                    b"number:text-style" => {
                        match e.name() {
                            b"number:boolean-style" => value_style.v_type = ValueType::Boolean,
                            b"number:date-style" => value_style.v_type = ValueType::DateTime,
                            b"number:time-style" => value_style.v_type = ValueType::TimeDuration,
                            b"number:number-style" => value_style.v_type = ValueType::Number,
                            b"number:currency-style" => value_style.v_type = ValueType::Currency,
                            b"number:percentage-style" => value_style.v_type = ValueType::Percentage,
                            b"number:text-style" => value_style.v_type = ValueType::Text,
                            _ => {}
                        }

                        for a in e.attributes().with_checks(false) {
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
                    }
                    b"number:boolean" => value_style.push_part(Part::new_prp(PartType::Boolean, attr_to_map(xml, e)?)),
                    b"number:number" => value_style.push_part(Part::new_prp(PartType::Number, attr_to_map(xml, e)?)),
                    b"number:scientific-number" => value_style.push_part(Part::new_prp(PartType::Scientific, attr_to_map(xml, e)?)),
                    b"number:day" => value_style.push_part(Part::new_prp(PartType::Day, attr_to_map(xml, e)?)),
                    b"number:month" => value_style.push_part(Part::new_prp(PartType::Month, attr_to_map(xml, e)?)),
                    b"number:year" => value_style.push_part(Part::new_prp(PartType::Year, attr_to_map(xml, e)?)),
                    b"number:era" => value_style.push_part(Part::new_prp(PartType::Era, attr_to_map(xml, e)?)),
                    b"number:day-of-week" => value_style.push_part(Part::new_prp(PartType::DayOfWeek, attr_to_map(xml, e)?)),
                    b"number:week-of-year" => value_style.push_part(Part::new_prp(PartType::WeekOfYear, attr_to_map(xml, e)?)),
                    b"number:quarter" => value_style.push_part(Part::new_prp(PartType::Quarter, attr_to_map(xml, e)?)),
                    b"number:hours" => value_style.push_part(Part::new_prp(PartType::Hours, attr_to_map(xml, e)?)),
                    b"number:minutes" => value_style.push_part(Part::new_prp(PartType::Minutes, attr_to_map(xml, e)?)),
                    b"number:seconds" => value_style.push_part(Part::new_prp(PartType::Seconds, attr_to_map(xml, e)?)),
                    b"number:fraction" => value_style.push_part(Part::new_prp(PartType::Fraction, attr_to_map(xml, e)?)),
                    b"number:am-pm" => value_style.push_part(Part::new_prp(PartType::AmPm, attr_to_map(xml, e)?)),
                    b"number:embedded-text" => value_style.push_part(Part::new_prp(PartType::EmbeddedText, attr_to_map(xml, e)?)),
                    b"number:text-content" => value_style.push_part(Part::new_prp(PartType::TextContent, attr_to_map(xml, e)?)),
                    b"style:text" => value_style.push_part(Part::new_prp(PartType::StyleText, attr_to_map(xml, e)?)),
                    b"style:map" => value_style.push_part(Part::new_prp(PartType::StyleMap, attr_to_map(xml, e)?)),

                    b"number:currency-symbol" => {
                        number_attr = Some(attr_to_map(xml, e)?);

                        if let Event::Empty(_) = evt {
                            let part = Part::new_prp(PartType::CurrencySymbol, number_attr.unwrap());
                            value_style.push_part(part);
                            number_attr = None;
                        }
                    }
                    b"number:text" => {
                        number_attr = Some(attr_to_map(xml, e)?);

                        if let Event::Empty(_) = evt {
                            let part = Part::new_prp(PartType::Text, number_attr.unwrap());
                            value_style.push_part(part);
                            number_attr = None;
                        }
                    }

                    _ => {}
                }
            }

            Event::Text(ref e) => {
                if DUMP_XML { log::debug!(" style {:?}", e); }

                let v = e.unescape_and_decode(&xml)?;
                cell_string = Some(v);
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
                    b"number:currency-symbol" => {
                        let mut part = Part::new(PartType::CurrencySymbol);
                        part.prp = number_attr;
                        part.content = cell_string;
                        value_style.push_part(part);

                        number_attr = None;
                        cell_string = None;
                    }
                    b"number:text" => {
                        let mut part = Part::new(PartType::Text);
                        part.prp = number_attr;
                        part.content = cell_string;
                        value_style.push_part(part);

                        number_attr = None;
                        cell_string = None;
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

fn attr_to_map(xml: &mut Reader<BufReader<&mut ZipFile>>, e: &BytesStart) -> Result<HashMap<String, String>, OdsError> {
    let mut m = HashMap::<String, String>::new();

    for a in e.attributes().with_checks(false) {
        if let Ok(ref attr) = a {
            let k = xml.decode(&attr.key)?;
            let v = attr.unescape_and_decode_value(&xml)?;
            m.insert(k.to_owned(), v);
        }
    }

    Ok(m)
}

/// Writes the ODS.
pub fn write_ods(book: &WorkBook, path: &Path) -> Result<(), OdsError> {
    let tmp_dir = tempfile::tempdir()?;

    if let Some(file) = &book.file {
        unzip(file, &tmp_dir)?;
    }

    let mut tmp_buf = tmp_dir.path().to_path_buf();
    tmp_buf.push("content.xml");
    write_ods_content(&tmp_buf, book)?;

    let mut tmp_buf = tmp_dir.path().to_path_buf();
    tmp_buf.push("META-INF/manifest.xml");
    if !tmp_buf.exists() {
        write_data(&tmp_buf, MANIFEST_XML_CONTENT)?;
    }

    create_zip(path, &tmp_dir)?;

    Ok(())
}

// Unzips all files to the tmpdir.
fn unzip(zip_file_path: &Path, tmpdir: &TempDir) -> Result<(), io::Error> {
    let zip_file = File::open(zip_file_path)?;
    // ods is a zip-archive, we read content.xml
    let mut zip = zip::ZipArchive::new(zip_file)?;

    for i in 0..zip.len() {
        let mut zip_entry = zip.by_index(i)?;

        if zip_entry.is_dir() {
            let mut out_dir_path = tmpdir.path().to_path_buf();
            out_dir_path.push(Path::new(zip_entry.name()));
            std::fs::create_dir_all(out_dir_path)?;
        } else {
            let mut out_file_path = tmpdir.path().to_path_buf();
            out_file_path.push(Path::new(zip_entry.name()));
            if let Some(parent) = out_file_path.parent() {
                std::fs::create_dir_all(parent)?;
            }
            let mut out_file = File::create(out_file_path)?;

            let mut buf: [u8; 1024] = [0; 1024];
            loop {
                let n = zip_entry.read(&mut buf)?;
                if n == 0 {
                    break;
                } else {
                    out_file.write_all(&buf[0..n])?;
                }
            }
        }
    }

    Ok(())
}

fn create_zip(zip_path: &Path, tmp_dir: &TempDir) -> Result<(), io::Error> {
    let zip_file = File::create(zip_path)?;
    let mut zip_writer = zip::ZipWriter::new(BufWriter::new(zip_file));
    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .unix_permissions(0o644);

    let walk = WalkDir::new(tmp_dir.path());
    for it in walk.into_iter().filter_map(|e| e.ok()) {
        let path = it.path();
        let name = path.strip_prefix(Path::new(tmp_dir.path())).unwrap().to_str().unwrap();

        if path.is_file() {
            zip_writer.start_file(name, options)?;

            let mut reader = File::open(path)?;
            let mut buf: [u8; 1024] = [0; 1024];
            loop {
                let n = reader.read(&mut buf)?;
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

fn write_ods_content(path: &Path, book: &WorkBook) -> Result<(), OdsError> {
    let f = File::create(path)?;

    let mut writer = Writer::new(BufWriter::new(f));

    writer.write_event(Event::Decl(BytesDecl::new(b"1.0", Some(b"UTF-8"), None)))?;
    writer.write(b"\n")?;
    writer.write_event(xml_start_a("office:document-content", vec![
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
    writer.write_event(xml_empty("office:scripts"))?;

    writer.write_event(xml_start("office:automatic-styles"))?;
    write_styles(&mut writer, &book.styles, Origin::Content)?;
    write_value_styles(&mut writer, &book.value_styles, Origin::Content)?;
    writer.write_event(xml_end("office:automatic-styles"))?;

    writer.write_event(xml_start("office:body"))?;
    writer.write_event(xml_start("office:spreadsheet"))?;

    for sheet in &book.sheets {
        write_sheet(&mut writer, &sheet, &book)?;
    }

    writer.write_event(xml_end("office:spreadsheet"))?;
    writer.write_event(xml_end("office:body"))?;
    writer.write_event(xml_end("office:document-content"))?;

    Ok(())
}

fn write_sheet(writer: &mut Writer<BufWriter<File>>, sheet: &Sheet, book: &WorkBook) -> Result<(), OdsError> {
    let mut attr: Vec<(&str, String)> = Vec::new();
    attr.push(("table-name", sheet.name.to_string()));
    if let Some(style) = &sheet.style {
        attr.push(("table:style", style.to_string()));
    }
    writer.write_event(xml_start_a("table:table", attr))?;

    let max_cell = sheet.used_grid_size();

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

        attr = Vec::new();
        if dc > 1 {
            attr.push(("table:number-columns-repeated", dc.to_string()));
        }
        if let Some(style) = &column.style {
            attr.push(("table:style-name", style.to_string()));
        }
        if let Some(style) = &column.def_cell_style {
            attr.push(("table:default-cell-style-name", style.to_string()));
        }
        writer.write_event(xml_empty_a("table:table-column", attr))?;
    }

    // table-row + table-cell
    let mut first_cell = true;
    let mut last_r: usize = 0;
    let mut last_c: usize = 0;

    for (r, c) in sheet.data.keys() {
        // For the repeat-counter we need to look forward.
        // Works nicely with the range operator :-)
        let (next_r, next_c) = if let Some(((next_r, next_c), _)) = sheet.data.range((*r, c + 1)..).next() {
            (*next_r, *next_c)
        } else {
            (max_cell.0, max_cell.1)
        };

        // Looking forward. Column deltas are only relevant in the same row.
        let forward_dr = next_r as i32 - *r as i32;
        // column advance only relevant in the same row
        let forward_dc = if forward_dr >= 1 {
            max_cell.1 as i32 - *c as i32
        } else {
            next_c as i32 - *c as i32
        };

        // Looking backward. Column deltas are only relevant in the same row.
        let backward_dr = *r as i32 - last_r as i32;
        let backward_dc = if backward_dr >= 1 {
            *c as i32
        } else {
            *c as i32 - last_c as i32
        };

        // log::info!("{} {} =={} {}==>  [[ {} {} ]] =={} {}==> {} {} ", last_r, last_c, backward_dr, backward_dc, *r, *c, forward_dr, forward_dc, next_r, next_c);

        // After the first cell there is always an open row tag.
        if backward_dr > 0 && !first_cell {
            writer.write_event(xml_end("table:table-row"))?;
        }

        // Any empty rows before this one?
        if backward_dr > 1 {
            // Empty rows in between are 1 less than the delta, except at the very start.
            attr = Vec::new();
            let empty_count = if first_cell { backward_dr } else { backward_dr - 1 };
            if empty_count > 1 {
                attr.push(("table:number-rows-repeated", empty_count.to_string()));
            }
            writer.write_event(xml_start_a("table:table-row", attr))?;
            // We fill the empty spaces completely up to max columns.
            writer.write_event(xml_empty_a("table:table-cell", vec![
                ("table:number-columns-repeated", max_cell.1.to_string()),
            ]))?;
            writer.write_event(xml_end("table:table-row"))?;
        }

        // Start a new row if there is a delta or we are at the start
        if backward_dr > 0 || first_cell {
            attr = Vec::new();
            if let Some(row) = sheet.rows.get(r) {
                if let Some(style) = &row.style {
                    attr.push(("table:style-name", style.to_string()));
                }
            }
            writer.write_event(xml_start_a("table:table-row", attr))?;

            // Might not be the first column in this row.
            if backward_dc > 0 {
                writer.write_event(xml_empty_a("table:table-cell", vec![
                    ("table:number-columns-repeated", backward_dc.to_string()),
                ]))?;
            }
        }

        // And now to something completely different ...
        let cell = &sheet.data[&(*r, *c)];

        attr = Vec::new();
        let mut content: String = String::from("");
        if let Some(formula) = &cell.formula {
            attr.push(("table:formula", formula.to_string()));
        }

        if let Some(style) = &cell.style {
            attr.push(("table:style-name", style.to_string()));
        } else {
            if let Some(value) = &cell.value {
                if let Some(style) = book.def_style(value.value_type()) {
                    attr.push(("table:style-name", style.to_string()));
                }
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
            writer.write_event(xml_start_a("table:table-cell", attr))?;
            writer.write_event(xml_start("text:p"))?;
            writer.write_event(xml_text(&content))?;
            writer.write_event(xml_end("text:p"))?;
            writer.write_event(xml_end("table:table-cell"))?;
        } else {
            writer.write_event(xml_empty_a("table:table-cell", attr))?;
        }

        // There may be some blank cells until the next one, but only one less the forward.
        if forward_dc > 1 {
            attr = Vec::new();
            attr.push(("table:number-columns-repeated", (forward_dc - 1).to_string()));
            writer.write_event(xml_empty_a("table:table-cell", attr))?;
        }

        first_cell = false;
        last_r = *r;
        last_c = *c;
    }

    if !first_cell {
        writer.write_event(xml_end("table:table-row"))?;
    }

    writer.write_event(xml_end("table:table"))?;

    Ok(())
}

fn write_styles(writer: &mut Writer<BufWriter<File>>, styles: &BTreeMap<String, Style>, origin: Origin) -> Result<(), OdsError> {
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
        writer.write_event(xml_start_a("style:style", attr))?;

        if let Some(table_cell_prp) = &style.table_cell_prp {
            writer.write_event(xml_empty_m("style:table-cell-properties", &table_cell_prp))?;
        }
        if let Some(table_col_prp) = &style.table_col_prp {
            writer.write_event(xml_empty_m("style:table-column-properties", &table_col_prp))?;
        }
        if let Some(table_row_prp) = &style.table_row_prp {
            writer.write_event(xml_empty_m("style:table-row-properties", &table_row_prp))?;
        }
        if let Some(table_prp) = &style.table_prp {
            writer.write_event(xml_empty_m("style:table-properties", &table_prp))?;
        }
        if let Some(paragraph_prp) = &style.paragraph_prp {
            writer.write_event(xml_empty_m("style:paragraph-properties", &paragraph_prp))?;
        }
        if let Some(text_prp) = &style.text_prp {
            writer.write_event(xml_empty_m("style:text-properties", &text_prp))?;
        }

        writer.write_event(xml_end("style:style"))?;
    }

    Ok(())
}

fn write_value_styles(writer: &mut Writer<BufWriter<File>>, styles: &BTreeMap<String, ValueStyle>, origin: Origin) -> Result<(), OdsError> {
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
        writer.write_event(xml_start_m(tag, &attr))?;

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
                    writer.write_event(xml_start_o(part_tag, part.prp.as_ref()))?;
                    if let Some(content) = &part.content {
                        writer.write_event(xml_text(content))?;
                    }
                    writer.write_event(xml_end(part_tag))?;
                } else {
                    writer.write_event(xml_empty_o(part_tag, part.prp.as_ref()))?;
                }
            }
        }
        writer.write_event(xml_end(tag))?;
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

const MANIFEST_XML_CONTENT: &'static str = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
 <manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>
 <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/>
 <manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"#;

fn write_data(out_file: &Path, data: &str) -> Result<(), io::Error> {
    if let Some(parent) = out_file.parent() {
        std::fs::create_dir_all(parent)?;
    }
    let mut f = File::create(out_file)?;
    f.write_all(data.as_bytes())?;

    Ok(())
}
