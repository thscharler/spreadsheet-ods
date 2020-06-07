//! Implements reading and writing of ODS Files.
//!
//! Warning ahead: This does'nt cover the full specification, just a
//! useable subset to read + modify + write back an ODS file.
//!
//! ```
//! use spreadsheet_ods::{WorkBook, Sheet};
//! use chrono::NaiveDate;
//! use spreadsheet_ods::format;
//! use spreadsheet_ods::formula;
//! use spreadsheet_ods::style::{Style, StyleFor};
//!
//! let mut wb = spreadsheet_ods::io::read_ods("tests/example.ods").unwrap();
//!
//! let mut sheet = wb.sheet_mut(0);
//! sheet.set_value(0, 0, 21.4f32);
//! sheet.set_value(0, 1, "foo");
//! sheet.set_styled_value(0, 2, NaiveDate::from_ymd(2020, 03, 01), "nice_date_style");
//! sheet.set_formula(0, 3, format!("of:={}+1", formula::fcellref(0,0)));
//!
//! let nice_date_format = format::create_date_dmy_format("nice_date_format");
//! wb.add_format(nice_date_format);
//!
//! let nice_date_style = Style::with_name(StyleFor::TableCell, "nice_date_style", "nice_date_format");
//! wb.add_style(nice_date_style);
//!
//! spreadsheet_ods::io::write_ods(&wb, "test_out/tryout.ods");
//!
//! ```
//!
//! When saving all the extra content is copied from the original file,
//! except for content.xml which is rewritten.
//!
//! For content.xml the following information is read and written:
//! TODO: more detailed description
//!
//! The following things are ignored for now
//! TODO: more detailed description
//!
//!

use std::collections::{BTreeMap, HashMap};
use std::fmt;
use std::path::PathBuf;

use chrono::{NaiveDate, NaiveDateTime};
use rust_decimal::Decimal;
use rust_decimal::prelude::*;
use time::Duration;

pub use error::OdsError;

use crate::format::ValueFormat;
use crate::refs::{CellRange, ColRange, RowRange};
use crate::style::{FontDecl, Style};

pub mod error;
pub mod io;
pub mod refs;
pub mod style;
pub mod format;
pub mod defaultstyles;
pub mod formula;
pub mod util;

/// Cell index type for row/column indexes.
#[allow(non_camel_case_types)]
pub type ucell = u32;

/// Origin of a style. Content.xml or Styles.xml.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum XMLOrigin {
    Content,
    Styles,
}

impl Default for XMLOrigin {
    fn default() -> Self {
        XMLOrigin::Content
    }
}

/// Book is the main structure for the Spreadsheet.
#[derive(Clone, Default)]
pub struct WorkBook {
    /// The data.
    sheets: Vec<Sheet>,

    //// FontDecl hold the style:font-face elements
    fonts: HashMap<String, FontDecl>,

    /// Styles hold the style:style elements.
    styles: HashMap<String, Style>,

    /// Value-styles are actual formatting instructions
    /// for various datatypes.
    /// Represents the various number:xxx-style elements.
    formats: HashMap<String, ValueFormat>,

    /// Default-styles per Type.
    /// This is only used when writing the ods file.
    def_styles: Option<HashMap<ValueType, String>>,

    /// TODO: Page layouts.
    // page_layouts: HashMap<String, PageLayout>,

    /// TODO: Page styles.
    // page_styles: HashMap<String, PageStyle>,

    /// Original file if this book was read from one.
    /// This is used when writing to copy all additional
    /// files except content.xml
    file: Option<PathBuf>,
}

impl fmt::Debug for WorkBook {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        for s in self.sheets.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.fonts.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.styles.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.def_styles.iter() {
            writeln!(f, "{:?}", s)?;
        }
        writeln!(f, "{:?}", self.file)?;
        Ok(())
    }
}

impl WorkBook {
    pub fn new() -> Self {
        WorkBook {
            sheets: Vec::new(),
            fonts: HashMap::new(),
            styles: HashMap::new(),
            formats: HashMap::new(),
            def_styles: None,
            file: None,
        }
    }

    /// Number of sheets.
    pub fn num_sheets(&self) -> usize {
        self.sheets.len()
    }

    /// Returns a certain sheet.
    /// panics if n does not exist.
    pub fn sheet(&self, n: usize) -> &Sheet {
        &self.sheets[n]
    }

    /// Returns a certain sheet.
    /// panics if n does not exist.
    pub fn sheet_mut(&mut self, n: usize) -> &mut Sheet {
        &mut self.sheets[n]
    }

    /// Inserts the sheet at the given position.
    pub fn insert_sheet(&mut self, i: usize, sheet: Sheet) {
        self.sheets.insert(i, sheet);
    }

    /// Appends a sheet.
    pub fn push_sheet(&mut self, sheet: Sheet) {
        self.sheets.push(sheet);
    }

    /// Removes a sheet from the table.
    pub fn remove_sheet(&mut self, n: usize) -> Sheet {
        self.sheets.remove(n)
    }

    /// Adds a default-style for all new values.
    /// This information is only used when writing the data to the ODS file.
    pub fn add_def_style(&mut self, value_type: ValueType, style: &str) {
        if self.def_styles.is_none() {
            self.def_styles = Some(HashMap::new());
        }

        if let Some(def_styles) = &mut self.def_styles {
            def_styles.insert(value_type, style.to_string());
        }
    }

    /// Returns the default style name.
    pub fn def_style(&self, value_type: ValueType) -> Option<&String> {
        if let Some(def_styles) = &self.def_styles {
            def_styles.get(&value_type)
        } else {
            None
        }
    }

    /// Finds a ValueFormat starting with the stylename attached to a cell.
    pub fn find_value_format(&self, style_name: &str) -> Option<&ValueFormat> {
        if let Some(style) = self.styles.get(style_name) {
            if let Some(value_format_name) = style.value_format() {
                if let Some(value_format) = self.formats.get(value_format_name) {
                    return Some(&value_format);
                }
            }
        }

        None
    }

    /// Adds a font.
    pub fn add_font(&mut self, font: FontDecl) {
        self.fonts.insert(font.name.to_string(), font);
    }

    /// Removes a font.
    pub fn remove_font(&mut self, name: &str) {
        self.fonts.remove(name);
    }

    /// Returns the FontDecl.
    pub fn font(&self, name: &str) -> Option<&FontDecl> {
        self.fonts.get(name)
    }

    /// Returns a mutable FontDecl.
    pub fn font_mut(&mut self, name: &str) -> Option<&mut FontDecl> {
        self.fonts.get_mut(name)
    }

    /// Adds a style.
    pub fn add_style(&mut self, style: Style) { self.styles.insert(style.name.to_string(), style); }

    /// Removes a style.
    pub fn remove_style(&mut self, name: &str) { self.styles.remove(name); }

    /// Returns the style.
    pub fn style(&self, name: &str) -> Option<&Style> { self.styles.get(name) }

    /// Returns the mutable style.
    pub fn style_mut(&mut self, name: &str) -> Option<&mut Style> {
        self.styles.get_mut(name)
    }

    /// Adds a value format.
    pub fn add_format(&mut self, vstyle: ValueFormat) {
        self.formats.insert(vstyle.name.to_string(), vstyle);
    }

    /// Removes the format.
    pub fn remove_format(&mut self, name: &str) {
        self.formats.remove(name);
    }

    /// Returns the format.
    pub fn format(&self, name: &str) -> Option<&ValueFormat> {
        self.formats.get(name)
    }

    /// Returns the mutable format.
    pub fn format_mut(&mut self, name: &str) -> Option<&mut ValueFormat> {
        self.formats.get_mut(name)
    }
}

/// One sheet of the spreadsheet.
///
/// Contains the data and the style-references. The can also be
/// styles on the whole sheet, columns and rows. The more complicated
/// grouping tags are not covered.
#[derive(Clone, Default)]
pub struct Sheet {
    name: String,
    style: Option<String>,

    data: BTreeMap<(ucell, ucell), SCell>,

    col_style: Option<BTreeMap<ucell, String>>,
    col_cell_style: Option<BTreeMap<ucell, String>>,
    row_style: Option<BTreeMap<ucell, String>>,

    header_rows: Option<RowRange>,
    header_cols: Option<ColRange>,
    print_ranges: Option<Vec<CellRange>>,
}

impl fmt::Debug for Sheet {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(f, "name {:?} style {:?}", self.name, self.style)?;
        for (k, v) in self.data.iter() {
            writeln!(f, "  data {:?} {:?}", k, v)?;
        }
        if let Some(col_style) = &self.col_style {
            for (k, v) in col_style {
                writeln!(f, "{:?} {:?}", k, v)?;
            }
        }
        if let Some(col_cell_style) = &self.col_cell_style {
            for (k, v) in col_cell_style {
                writeln!(f, "{:?} {:?}", k, v)?;
            }
        }
        if let Some(row_style) = &self.row_style {
            for (k, v) in row_style {
                writeln!(f, "{:?} {:?}", k, v)?;
            }
        }
        if let Some(header_rows) = &self.header_rows {
            writeln!(f, "header rows {:?}", header_rows)?;
        }
        if let Some(header_cols) = &self.header_cols {
            writeln!(f, "header cols {:?}", header_cols)?;
        }
        Ok(())
    }
}


impl Sheet {
    /// New, empty
    pub fn new() -> Self {
        Sheet {
            name: String::from(""),
            data: BTreeMap::new(),
            style: None,
            col_style: None,
            col_cell_style: None,
            row_style: None,
            header_rows: None,
            header_cols: None,
            print_ranges: None,

        }
    }

    // New, empty, but with a name.
    pub fn with_name<S: Into<String>>(name: S) -> Self {
        Sheet {
            name: name.into(),
            data: BTreeMap::new(),
            style: None,
            col_style: None,
            col_cell_style: None,
            row_style: None,
            header_rows: None,
            header_cols: None,
            print_ranges: None,

        }
    }

    /// Sheet name.
    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    /// Sheet name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets the table-style
    pub fn set_style<V: Into<String>>(&mut self, style: V) {
        self.style = Some(style.into());
    }

    /// Returns the table-style.
    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    /// Column wide style.
    pub fn set_column_style<V: Into<String>>(&mut self, col: ucell, style: V) {
        if self.col_style.is_none() {
            self.col_style = Some(BTreeMap::new());
        }
        if let Some(col_style) = &mut self.col_style {
            col_style.entry(col).or_insert_with(|| style.into());
        }
    }

    /// Returns the column wide style.
    pub fn column_style(&self, col: ucell) -> Option<&String> {
        if let Some(col_style) = &self.col_style {
            col_style.get(&col)
        } else {
            None
        }
    }

    /// Default cell style for this column.
    pub fn set_column_cell_style<V: Into<String>>(&mut self, col: ucell, style: V) {
        if self.col_cell_style.is_none() {
            self.col_cell_style = Some(BTreeMap::new());
        }
        if let Some(col_cell_style) = &mut self.col_cell_style {
            col_cell_style.entry(col).or_insert_with(|| style.into());
        }
    }

    /// Returns the default cell style for this column.
    pub fn column_cell_style(&self, col: ucell) -> Option<&String> {
        if let Some(col_cell_style) = &self.col_cell_style {
            col_cell_style.get(&col)
        } else {
            None
        }
    }

    /// Row style.
    pub fn set_row_style<V: Into<String>>(&mut self, row: ucell, style: V) {
        if self.row_style.is_none() {
            self.row_style = Some(BTreeMap::new());
        }
        if let Some(row_style) = &mut self.row_style {
            row_style.entry(row).or_insert_with(|| style.into());
        }
    }

    /// Returns the row style.
    pub fn row_style(&self, row: ucell) -> Option<&String> {
        if let Some(row_style) = &self.row_style {
            row_style.get(&row)
        } else {
            None
        }
    }

    /// Returns a tuple of (max(row)+1, max(col)+1)
    pub fn used_grid_size(&self) -> (ucell, ucell) {
        let max = self.data.keys().fold((0, 0), |mut max, (r, c)| {
            max.0 = u32::max(max.0, *r);
            max.1 = u32::max(max.1, *c);
            max
        });

        (max.0 + 1, max.1 + 1)
    }

    /// Returns the cell if available.
    pub fn cell(&self, row: ucell, col: ucell) -> Option<&SCell> {
        self.data.get(&(row, col))
    }

    /// Returns a mutable reference to the cell.
    pub fn cell_mut(&mut self, row: ucell, col: ucell) -> Option<&mut SCell> {
        self.data.get_mut(&(row, col))
    }

    /// Creates an empty cell if the position is currently empty and returns
    /// a reference.
    pub fn create_cell(&mut self, row: ucell, col: ucell) -> &mut SCell {
        self.data.entry((row, col)).or_insert_with(SCell::new)
    }

    /// Adds a cell. Replaces an existing one.
    pub fn add_cell(&mut self, row: ucell, col: ucell, cell: SCell) -> Option<SCell> {
        self.data.insert((row, col), cell)
    }

    // Removes a value.
    pub fn remove_cell(&mut self, row: ucell, col: ucell) -> Option<SCell> {
        self.data.remove(&(row, col))
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_styled_value<V: Into<Value>, W: Into<String>>(&mut self, row: ucell, col: ucell, value: V, style: W) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.value = value.into();
        cell.style = Some(style.into());
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_value<V: Into<Value>>(&mut self, row: ucell, col: ucell, value: V) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.value = value.into();
    }

    /// Returns a value
    pub fn value(&self, row: ucell, col: ucell) -> &Value {
        if let Some(cell) = self.data.get(&(row, col)) {
            &cell.value
        } else {
            &Value::Empty
        }
    }

    /// Sets a formula for the specified cell. Creates a new cell if necessary.
    pub fn set_formula<V: Into<String>>(&mut self, row: ucell, col: ucell, formula: V) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.formula = Some(formula.into());
    }

    /// Returns a value
    pub fn formula(&self, row: ucell, col: ucell) -> Option<&String> {
        if let Some(c) = self.data.get(&(row, col)) {
            c.formula.as_ref()
        } else {
            None
        }
    }

    /// Sets the cell-style for the specified cell. Creates a new cell if necessary.
    pub fn set_cell_style<V: Into<String>>(&mut self, row: ucell, col: ucell, style: V) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.style = Some(style.into());
    }

    /// Returns a value
    pub fn cell_style(&self, row: ucell, col: ucell) -> Option<&String> {
        if let Some(c) = self.data.get(&(row, col)) {
            c.style.as_ref()
        } else {
            None
        }
    }

    /// Sets the rowspan of the cell. Must be greater than 0.
    pub fn set_row_span(&mut self, row: ucell, col: ucell, span: ucell) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.span.0 = span;
    }

    // Rowspan of the cell.
    pub fn row_span(&self, row: ucell, col: ucell) -> ucell {
        if let Some(c) = self.data.get(&(row, col)) {
            c.span.0
        } else {
            1
        }
    }

    /// Sets the colspan of the cell. Must be greater than 0.
    pub fn set_col_span(&mut self, row: ucell, col: ucell, span: ucell) {
        assert!(span > 0);
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.span.1 = span;
    }

    /// Colspan of the cell.
    pub fn col_span(&self, row: ucell, col: ucell) -> ucell {
        if let Some(c) = self.data.get(&(row, col)) {
            c.span.1
        } else {
            1
        }
    }

    /// Defines a range of rows as header rows.
    pub fn set_header_rows(&mut self, row_start: ucell, row_end: ucell) {
        self.header_rows = Some(RowRange::new(row_start, row_end));
    }

    /// Clears the header-rows definition.
    pub fn clear_header_rows(&mut self) {
        self.header_rows = None;
    }

    /// Returns the header rows.
    pub fn header_rows(&self) -> &Option<RowRange> {
        &self.header_rows
    }

    /// Defines a range of columns as header columns.
    pub fn set_header_cols(&mut self, col_start: ucell, col_end: ucell) {
        self.header_cols = Some(ColRange::new(col_start, col_end));
    }

    /// Clears the header-columns definition.
    pub fn clear_header_cols(&mut self) {
        self.header_cols = None;
    }

    /// Returns the header columns.
    pub fn header_cols(&self) -> &Option<ColRange> {
        &self.header_cols
    }

    /// Print ranges.
    pub fn add_print_range(&mut self, range: CellRange) {
        if self.print_ranges.is_none() {
            self.print_ranges = Some(Vec::new());
        }
        if let Some(ref mut print_ranges) = self.print_ranges {
            print_ranges.push(range);
        }
    }

    /// Remove print ranges.
    pub fn clear_print_ranges(&mut self) {
        self.print_ranges = None;
    }

    /// Return the print ranges.
    pub fn print_ranges(&self) -> &Option<Vec<CellRange>> {
        &self.print_ranges
    }
}

/// One Cell of the spreadsheet.
#[derive(Debug, Clone, Default)]
pub struct SCell {
    value: Value,
    // Unparsed formula string.
    formula: Option<String>,
    // Cell style name.
    style: Option<String>,
    // Row/Column span.
    span: (ucell, ucell),
}

impl SCell {
    /// New, empty.
    pub fn new() -> Self {
        SCell {
            value: Value::Empty,
            formula: None,
            style: None,
            span: (1, 1),
        }
    }

    /// New, with a value.
    pub fn with_value<V: Into<Value>>(value: V) -> Self {
        SCell {
            value: value.into(),
            formula: None,
            style: None,
            span: (1, 1),
        }
    }

    /// Returns the value.
    pub fn value(&self) -> &Value {
        &self.value
    }

    /// Sets the value.
    pub fn set_value<V: Into<Value>>(&mut self, value: V) {
        self.value = value.into();
    }

    /// Returns the formula.
    pub fn formula(&self) -> Option<&String> {
        self.formula.as_ref()
    }

    /// Sets the formula.
    pub fn set_formula<V: Into<String>>(&mut self, formula: V) {
        self.formula = Some(formula.into());
    }

    /// Returns the cell style.
    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    /// Sets the cell style.
    pub fn set_style<V: Into<String>>(&mut self, style: V) {
        self.style = Some(style.into());
    }

    /// Sets the row span of this cell.
    /// Cells below with values will be lost when writing.
    pub fn set_row_span(&mut self, rows: ucell) {
        assert!(rows > 0);
        self.span.0 = rows;
    }

    /// Returns the row span.
    pub fn row_span(&self) -> ucell {
        self.span.0
    }

    /// Sets the column span of this cell.
    /// Cells to the right with values will be lost when writing.
    pub fn set_col_span(&mut self, cols: ucell) {
        assert!(cols > 0);
        self.span.1 = cols;
    }

    /// Returns the col span.
    pub fn col_span(&self) -> ucell {
        self.span.1
    }
}

/// Datatypes
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ValueType {
    Empty,
    Boolean,
    Number,
    Percentage,
    Currency,
    Text,
    DateTime,
    TimeDuration,
}

impl Default for ValueType {
    fn default() -> Self {
        ValueType::Text
    }
}

/// Content-Values
#[derive(Debug, Clone)]
pub enum Value {
    Empty,
    Boolean(bool),
    Number(f64),
    Percentage(f64),
    Currency(String, f64),
    Text(String),
    // TODO: CompositText(...)
    // TODO: Frame(Image)
    DateTime(NaiveDateTime),
    TimeDuration(Duration),
}

impl Value {
    pub fn value_type(&self) -> ValueType {
        match self {
            Value::Empty => ValueType::Empty,
            Value::Boolean(_) => ValueType::Boolean,
            Value::Number(_) => ValueType::Number,
            Value::Percentage(_) => ValueType::Percentage,
            Value::Currency(_, _) => ValueType::Currency,
            Value::Text(_) => ValueType::Text,
            Value::TimeDuration(_) => ValueType::TimeDuration,
            Value::DateTime(_) => ValueType::DateTime,
        }
    }

    /// Return content.
    pub fn as_bool_or(&self, d: bool) -> bool {
        match self {
            Value::Boolean(b) => *b,
            _ => d,
        }
    }

    pub fn as_i32_or(&self, d: i32) -> i32 {
        match self {
            Value::Number(n) => *n as i32,
            Value::Percentage(p) => *p as i32,
            Value::Currency(_, v) => *v as i32,
            _ => d,
        }
    }

    pub fn as_u32_or(&self, d: u32) -> u32 {
        match self {
            Value::Number(n) => *n as u32,
            Value::Percentage(p) => *p as u32,
            Value::Currency(_, v) => *v as u32,
            _ => d,
        }
    }

    pub fn as_decimal_or(&self, d: Decimal) -> Decimal {
        match self {
            Value::Number(n) => Decimal::from_f64(*n).unwrap(),
            Value::Currency(_, v) => Decimal::from_f64(*v).unwrap(),
            Value::Percentage(p) => Decimal::from_f64(*p).unwrap(),
            _ => d,
        }
    }

    pub fn as_f64_or(&self, d: f64) -> f64 {
        match self {
            Value::Number(n) => *n,
            Value::Currency(_, v) => *v,
            Value::Percentage(p) => *p,
            _ => d,
        }
    }

    pub fn as_str_or<'a>(&'a self, d: &'a str) -> &'a str {
        match self {
            Value::Text(s) => s.as_ref(),
            _ => d,
        }
    }

    pub fn as_timeduration_or(&self, d: Duration) -> Duration {
        match self {
            Value::TimeDuration(td) => *td,
            _ => d,
        }
    }

    pub fn as_datetime_or(&self, d: NaiveDateTime) -> NaiveDateTime {
        match self {
            Value::DateTime(dt) => *dt,
            _ => d,
        }
    }

    pub fn as_datetime_opt(&self) -> Option<NaiveDateTime> {
        match self {
            Value::DateTime(dt) => Some(*dt),
            _ => None,
        }
    }

    pub fn currency(currency: &str, value: f64) -> Self {
        Value::Currency(currency.to_string(), value)
    }

    pub fn percentage(percent: f64) -> Self {
        Value::Percentage(percent)
    }
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}

impl From<&str> for Value {
    fn from(s: &str) -> Self {
        Value::Text(s.to_string())
    }
}

impl From<String> for Value {
    fn from(s: String) -> Self {
        Value::Text(s)
    }
}

impl From<&String> for Value {
    fn from(s: &String) -> Self {
        Value::Text(s.to_string())
    }
}

impl From<Option<String>> for Value {
    fn from(s: Option<String>) -> Self {
        if let Some(s) = s {
            Value::Text(s)
        } else {
            Value::Empty
        }
    }
}

impl From<Decimal> for Value {
    fn from(f: Decimal) -> Self {
        Value::Number(f.to_f64().unwrap())
    }
}

impl From<f64> for Value {
    fn from(f: f64) -> Self {
        Value::Number(f)
    }
}

impl From<f32> for Value {
    fn from(f: f32) -> Self { Value::Number(f as f64) }
}

impl From<i64> for Value {
    fn from(i: i64) -> Self {
        Value::Number(i as f64)
    }
}

impl From<i32> for Value {
    fn from(i: i32) -> Self {
        Value::Number(i as f64)
    }
}

impl From<u64> for Value {
    fn from(u: u64) -> Self {
        Value::Number(u as f64)
    }
}

impl From<u32> for Value {
    fn from(u: u32) -> Self {
        Value::Number(u as f64)
    }
}

impl From<bool> for Value {
    fn from(b: bool) -> Self {
        Value::Boolean(b)
    }
}

impl From<NaiveDateTime> for Value {
    fn from(dt: NaiveDateTime) -> Self {
        Value::DateTime(dt)
    }
}

impl From<Option<NaiveDateTime>> for Value {
    fn from(dt: Option<NaiveDateTime>) -> Self {
        if let Some(dt) = dt {
            Value::DateTime(dt)
        } else {
            Value::Empty
        }
    }
}

impl From<NaiveDate> for Value {
    fn from(dt: NaiveDate) -> Self { Value::DateTime(dt.and_hms(0, 0, 0)) }
}

impl From<Option<NaiveDate>> for Value {
    fn from(dt: Option<NaiveDate>) -> Self {
        if let Some(dt) = dt {
            Value::DateTime(dt.and_hms(0, 0, 0))
        } else {
            Value::Empty
        }
    }
}

impl From<Duration> for Value {
    fn from(d: Duration) -> Self {
        Value::TimeDuration(d)
    }
}
