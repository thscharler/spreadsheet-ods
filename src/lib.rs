//! Implements reading and writing of ODS Files.
//!
//! Warning ahead: This does'nt cover the full specification, just a
//! useable subset to read + modify + write back an ODS file.
//!
//! ```
//! use spreadsheet_ods::{WorkBook, Sheet, Style, StyleFor, ValueFormat, ValueType, FormatPart};
//! use chrono::NaiveDate;
//! use spreadsheet_ods::format;
//! use spreadsheet_ods::formula;
//!
//! let mut wb = spreadsheet_ods::ods::read_ods("tests/example.ods").unwrap();
//!
//! let mut sheet = wb.sheet_mut(0);
//! sheet.set_value(0, 0, 21.4f32);
//! sheet.set_value(0, 1, "foo");
//! sheet.set_styled_value(0, 2, NaiveDate::from_ymd(2020, 03, 01), "nice_date_style");
//! sheet.set_formula(0, 3, format!("of:={}+1", formula::cellref(0,0)));
//!
//! let nice_date_format = format::create_date_dmy_format("nice_date_format");
//! wb.add_format(nice_date_format);
//!
//! let nice_date_style = Style::with_name(StyleFor::TableCell, "nice_date_style", "nice_date_format");
//! wb.add_style(nice_date_style);
//!
//! spreadsheet_ods::ods::write_ods(&wb, "tryout.ods");
//!
//! ```
//!
//! When saving all the extra content is copied from the original file,
//! except for content.xml which is rewritten.
//!
//! For content.xml the following information is read and written:
//! * fonts
//! * styles
//! * data-formats
//! * table-data
//! ** all datatypes
//! ** no complex text
//!
//! The following things are ignored for now
//! * conditional formats
//! * charts
//! * ...
//!
//!

use std::collections::{BTreeMap, HashMap};
use std::fmt;
use std::fmt::{Display, Formatter};
use std::path::PathBuf;

use chrono::{Duration, NaiveDate, NaiveDateTime};
use rust_decimal::Decimal;
use rust_decimal::prelude::*;
use string_cache::DefaultAtom;

pub mod ods;
pub mod style;
pub mod defaultstyles;
pub mod format;
pub mod formula;

/// Cell index type for row/column indexes.
#[allow(non_camel_case_types)]
pub type ucell = u32;

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

    /// Page layouts.
    // page_layouts: HashMap<String, PageLayout>,

    /// Page styles.
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

    // Colspan of the cell.
    pub fn col_span(&self, row: ucell, col: ucell) -> ucell {
        if let Some(c) = self.data.get(&(row, col)) {
            c.span.1
        } else {
            1
        }
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

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontDecl {
    name: String,
    /// From where did we get this style.
    origin: XMLOrigin,
    /// All other attributes.
    prp: Option<HashMap<DefaultAtom, String>>,
}

impl FontDecl {
    /// New, empty.
    pub fn new() -> Self {
        FontDecl::new_origin(XMLOrigin::Content)
    }

    /// New, with origination.
    pub fn new_origin(origin: XMLOrigin) -> Self {
        Self {
            name: "".to_string(),
            origin,
            prp: None,
        }
    }

    /// New, with a name.
    pub fn with_name<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            origin: XMLOrigin::Content,
            prp: None,
        }
    }

    /// Set the name.
    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets a property of the font.
    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    /// Returns a property of the font.
    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }
}

/// Style data fashioned after the ODS spec.
#[derive(Debug, Clone, Default)]
pub struct Style {
    name: String,
    /// Nice String.
    display_name: Option<String>,
    /// From where did we get this style.
    origin: XMLOrigin,
    /// Applicability of this style.
    family: StyleFor,
    /// Styles can cascade.
    parent: Option<String>,
    /// References the actual formatting instructions in the value-styles.
    value_format: Option<String>,
    /// Table styling
    table_prp: Option<HashMap<DefaultAtom, String>>,
    /// Column styling
    table_col_prp: Option<HashMap<DefaultAtom, String>>,
    /// Row styling
    table_row_prp: Option<HashMap<DefaultAtom, String>>,
    /// Cell styles
    table_cell_prp: Option<HashMap<DefaultAtom, String>>,
    /// Cell paragraph styles
    paragraph_prp: Option<HashMap<DefaultAtom, String>>,
    /// Cell text styles
    text_prp: Option<HashMap<DefaultAtom, String>>,
}

impl Style {
    /// New, empty.
    pub fn new() -> Self {
        Style::new_origin(XMLOrigin::Content)
    }

    /// New, with origination.
    pub fn new_origin(origin: XMLOrigin) -> Self {
        Style {
            name: String::from(""),
            display_name: None,
            origin,
            family: StyleFor::None,
            parent: None,
            value_format: None,
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    pub fn cell_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableCell, name, value_style)
    }

    pub fn col_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableColumn, name, value_style)
    }

    pub fn row_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableRow, name, value_style)
    }

    pub fn table_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::Table, name, value_style)
    }

    /// New, with name.
    pub fn with_name<S: Into<String>, T: Into<String>>(family: StyleFor, name: S, value_style: T) -> Self {
        Style {
            name: name.into(),
            display_name: None,
            origin: XMLOrigin::Content,
            family,
            parent: Some(String::from("Default")),
            value_format: Some(value_style.into()),
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    /// Sets the name.
    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets the display name.
    pub fn set_display_name(&mut self, name: &str) {
        self.display_name = Some(name.to_string());
    }

    /// Returns the display name.
    pub fn display_name(&self) -> Option<&String> {
        self.display_name.as_ref()
    }

    /// Sets the origin.
    pub fn set_origin(&mut self, origin: XMLOrigin) {
        self.origin = origin;
    }

    /// Returns the origin.
    pub fn origin(&self) -> &XMLOrigin {
        &self.origin
    }

    /// Sets the style-family.
    pub fn set_family(&mut self, family: StyleFor) {
        self.family = family;
    }

    /// Returns the style-family.
    pub fn family(&self) -> &StyleFor {
        &self.family
    }

    /// Sets the parent style.
    pub fn set_parent(&mut self, parent: &str) {
        self.parent = Some(parent.to_string());
    }

    /// Returns the parent style.
    pub fn parent(&self) -> Option<&String> {
        self.parent.as_ref()
    }

    /// Sets the value format.
    pub fn set_value_format(&mut self, value_format: &str) {
        self.value_format = Some(value_format.to_string());
    }

    /// Returns the value format.
    pub fn value_format(&self) -> Option<&String> {
        self.value_format.as_ref()
    }

    /// Sets a property for a table style.
    pub fn set_table_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_prp, name, value);
    }

    /// Returns a property for a table style.
    pub fn table_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_prp, name)
    }

    /// Sets a property for a table column.
    pub fn set_table_col_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_col_prp, name, value);
    }

    /// Returns a property for a table column.
    pub fn table_col_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_col_prp, name)
    }

    /// Set a table row property.
    pub fn set_table_row_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_row_prp, name, value);
    }

    /// Returns a table row property.
    pub fn table_row_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_row_prp, name)
    }

    /// Sets a table cell property.
    pub fn set_table_cell_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_cell_prp, name, value);
    }

    /// Returns a table cell property.
    pub fn table_cell_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_cell_prp, name)
    }

    /// Sets a text property.
    pub fn set_text_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.text_prp, name, value);
    }

    /// Removes a text property.
    pub fn clear_text_prp(&mut self, name: &str) -> Option<String> {
        clear_prp(&mut self.text_prp, name)
    }

    /// Returns a text property.
    pub fn text_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.text_prp, name)
    }

    /// Sets a paragraph property.
    pub fn set_paragraph_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.paragraph_prp, name, value);
    }

    /// Returns a paragraph property.
    pub fn paragraph_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.paragraph_prp, name)
    }
}

/// Origin of a style. Content.xml or Styles.xml.
#[derive(Debug, Clone, PartialEq)]
pub enum XMLOrigin {
    Content,
    Styles,
}

impl Default for XMLOrigin {
    fn default() -> Self {
        XMLOrigin::Content
    }
}

/// Applicability of this style.
#[derive(Debug, Clone, PartialEq)]
pub enum StyleFor {
    Table,
    TableRow,
    TableColumn,
    TableCell,
    None,
}

impl Default for StyleFor {
    fn default() -> Self {
        StyleFor::None
    }
}

#[derive(Debug)]
pub enum ValueFormatError {
    Format(String),
    NaN,
}

impl Display for ValueFormatError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            ValueFormatError::Format(s) => write!(f, "{}", s)?,
            ValueFormatError::NaN => write!(f, "Digit expected")?,
        }
        Ok(())
    }
}

impl std::error::Error for ValueFormatError {}

/// Actual textual formatting of values.
#[derive(Debug, Clone, Default)]
pub struct ValueFormat {
    // Name
    name: String,
    // Value type
    v_type: ValueType,
    // Origin information.
    origin: XMLOrigin,
    // Properties of the format.
    prp: Option<HashMap<DefaultAtom, String>>,
    // Parts of the format.
    parts: Option<Vec<FormatPart>>,
}

impl ValueFormat {
    /// New, empty.
    pub fn new() -> Self {
        ValueFormat::new_origin(XMLOrigin::Content)
    }

    /// New, with origin.
    pub fn new_origin(origin: XMLOrigin) -> Self {
        ValueFormat {
            name: String::from(""),
            v_type: ValueType::Text,
            origin,
            prp: None,
            parts: None,
        }
    }

    /// New, with name.
    pub fn with_name<S: Into<String>>(name: S, value_type: ValueType) -> Self {
        ValueFormat {
            name: name.into(),
            v_type: value_type,
            origin: XMLOrigin::Content,
            prp: None,
            parts: None,
        }
    }

    /// Sets the name.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets the value type.
    pub fn set_value_type(&mut self, value_type: ValueType) {
        self.v_type = value_type;
    }

    /// Returns the value type.
    pub fn value_type(&self) -> &ValueType {
        &self.v_type
    }

    /// Sets the origin.
    pub fn set_origin(&mut self, origin: XMLOrigin) {
        self.origin = origin;
    }

    /// Returns the origin.
    pub fn origin(&self) -> &XMLOrigin {
        &self.origin
    }

    /// Sets a property of the format.
    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    /// Returns a property of the format.
    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    /// Adds a format part.
    pub fn push_part(&mut self, part: FormatPart) {
        if let Some(parts) = &mut self.parts {
            parts.push(part);
        } else {
            self.parts = Some(vec![part]);
        }
    }

    /// Adds all format parts.
    pub fn push_parts(&mut self, parts: Vec<FormatPart>) {
        for p in parts.into_iter() {
            self.push_part(p);
        }
    }

    /// Returns the parts.
    pub fn parts(&self) -> Option<&Vec<FormatPart>> {
        self.parts.as_ref()
    }

    /// Returns the mutable parts.
    pub fn parts_mut(&mut self) -> &mut Vec<FormatPart> {
        self.parts.get_or_insert(Vec::new())
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    pub fn format_boolean(&self, b: bool) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_boolean(&mut buf, b);
            }
        }
        buf
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    pub fn format_float(&self, f: f64) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_float(&mut buf, f);
            }
        }
        buf
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    pub fn format_str(&self, s: &str) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_str(&mut buf, s);
            }
        }
        buf
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    // Should work reasonably. Don't ask me about other calenders.
    pub fn format_datetime(&self, d: &NaiveDateTime) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            let h12 = parts.iter().any(|v| v.part_type == FormatPartType::AmPm);

            for p in parts {
                p.format_datetime(&mut buf, d, h12);
            }
        }
        buf
    }

    // Tries to format. Should work reasonably.
    // If there are no matching parts, does nothing.
    pub fn format_time_duration(&self, d: &Duration) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_time_duration(&mut buf, d);
            }
        }
        buf
    }
}

/// Identifies the structural parts of a value format.
#[derive(Debug, Clone, PartialEq)]
pub enum FormatPartType {
    Boolean,
    Number,
    Fraction,
    Scientific,
    CurrencySymbol,
    Day,
    Month,
    Year,
    Era,
    DayOfWeek,
    WeekOfYear,
    Quarter,
    Hours,
    Minutes,
    Seconds,
    AmPm,
    EmbeddedText,
    Text,
    TextContent,
    StyleText,
    StyleMap,
}

/// One structural part of a value format.
#[derive(Debug, Clone)]
pub struct FormatPart {
    // What kind of format part is this?
    part_type: FormatPartType,
    // Properties of this part.
    prp: Option<HashMap<DefaultAtom, String>>,
    // Some content.
    content: Option<String>,
}

impl FormatPart {
    /// New, empty
    pub fn new(ftype: FormatPartType) -> Self {
        FormatPart {
            part_type: ftype,
            prp: None,
            content: None,
        }
    }

    /// New, with string content.
    pub fn new_content(ftype: FormatPartType, content: &str) -> Self {
        FormatPart {
            part_type: ftype,
            prp: None,
            content: Some(content.to_string()),
        }
    }

    /// New with properties.
    pub fn new_vec(ftype: FormatPartType, prp_vec: Vec<(&str, String)>) -> Self {
        let mut part = FormatPart {
            part_type: ftype,
            prp: None,
            content: None,
        };
        part.set_prp_vec(prp_vec);
        part
    }

    /// Sets the kind of the part.
    pub fn set_part_type(&mut self, p_type: FormatPartType) {
        self.part_type = p_type;
    }

    /// What kind of part?
    pub fn part_type(&self) -> &FormatPartType {
        &self.part_type
    }

    /// Sets a vec of properties.
    pub fn set_prp_vec(&mut self, vec: Vec<(&str, String)>) {
        set_prp_vec(&mut self.prp, vec);
    }

    /// Sets a property.
    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    /// Returns a property.
    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    /// Returns a property or a default.
    pub fn prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.prp, name, default)
    }

    /// Sets a textual content for this part. This is only used
    /// for text and currency-symbol.
    pub fn set_content(&mut self, content: &str) {
        self.content = Some(content.to_string());
    }

    /// Returns the text content.
    pub fn content(&self) -> Option<&String> {
        self.content.as_ref()
    }

    /// Tries to format the given boolean, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_boolean(&self, buf: &mut String, b: bool) {
        match self.part_type {
            FormatPartType::Boolean => {
                buf.push_str(if b { "true" } else { "false" });
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given float, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_float(&self, buf: &mut String, f: f64) {
        match self.part_type {
            FormatPartType::Number => {
                let dec = self.prp_def("number:decimal-places", "0").parse::<usize>();
                if let Ok(dec) = dec {
                    buf.push_str(&format!("{:.*}", dec, f));
                }
            }
            FormatPartType::Scientific => {
                buf.push_str(&format!("{:e}", f));
            }
            FormatPartType::CurrencySymbol => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given string, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_str(&self, buf: &mut String, s: &str) {
        match self.part_type {
            FormatPartType::TextContent => {
                buf.push_str(s);
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given DateTime, and appends the result to buf.
    /// Uses chrono::strftime for the implementation.
    /// If this part does'nt match does nothing
    #[allow(clippy::collapsible_if)]
    fn format_datetime(&self, buf: &mut String, d: &NaiveDateTime, h12: bool) {
        match self.part_type {
            FormatPartType::Day => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%d").to_string());
                } else {
                    buf.push_str(&d.format("%-d").to_string());
                }
            }
            FormatPartType::Month => {
                let is_long = self.prp_def("number:style", "") == "long";
                let is_text = self.prp_def("number:textual", "") == "true";
                if is_text {
                    if is_long {
                        buf.push_str(&d.format("%b").to_string());
                    } else {
                        buf.push_str(&d.format("%B").to_string());
                    }
                } else {
                    if is_long {
                        buf.push_str(&d.format("%m").to_string());
                    } else {
                        buf.push_str(&d.format("%-m").to_string());
                    }
                }
            }
            FormatPartType::Year => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%Y").to_string());
                } else {
                    buf.push_str(&d.format("%y").to_string());
                }
            }
            FormatPartType::DayOfWeek => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%A").to_string());
                } else {
                    buf.push_str(&d.format("%a").to_string());
                }
            }
            FormatPartType::WeekOfYear => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%W").to_string());
                } else {
                    buf.push_str(&d.format("%-W").to_string());
                }
            }
            FormatPartType::Hours => {
                let is_long = self.prp_def("number:style", "") == "long";
                if !h12 {
                    if is_long {
                        buf.push_str(&d.format("%H").to_string());
                    } else {
                        buf.push_str(&d.format("%-H").to_string());
                    }
                } else {
                    if is_long {
                        buf.push_str(&d.format("%I").to_string());
                    } else {
                        buf.push_str(&d.format("%-I").to_string());
                    }
                }
            }
            FormatPartType::Minutes => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%M").to_string());
                } else {
                    buf.push_str(&d.format("%-M").to_string());
                }
            }
            FormatPartType::Seconds => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%S").to_string());
                } else {
                    buf.push_str(&d.format("%-S").to_string());
                }
            }
            FormatPartType::AmPm => {
                buf.push_str(&d.format("%p").to_string());
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given Duration, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_time_duration(&self, buf: &mut String, d: &Duration) {
        match self.part_type {
            FormatPartType::Hours => {
                buf.push_str(&d.num_hours().to_string());
            }
            FormatPartType::Minutes => {
                buf.push_str(&(d.num_minutes() % 60).to_string());
            }
            FormatPartType::Seconds => {
                buf.push_str(&(d.num_seconds() % 60).to_string());
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }
}

fn set_prp_vec(map: &mut Option<HashMap<DefaultAtom, String>>, vec: Vec<(&str, String)>) {
    if map.is_none() {
        map.replace(HashMap::new());
    }
    if let Some(map) = map {
        for (name, value) in vec {
            let a = DefaultAtom::from(name);
            map.insert(a, value);
        }
    }
}

fn set_prp(map: &mut Option<HashMap<DefaultAtom, String>>, name: &str, value: String) {
    if map.is_none() {
        map.replace(HashMap::new());
    }
    if let Some(map) = map {
        let a = DefaultAtom::from(name);
        map.insert(a, value);
    }
}

fn clear_prp(map: &mut Option<HashMap<DefaultAtom, String>>, name: &str) -> Option<String> {
    if !map.is_none() {
        if let Some(map) = map {
            map.remove(&DefaultAtom::from(name))
        } else {
            None
        }
    } else {
        None
    }
}

fn get_prp<'a, 'b>(map: &'a Option<HashMap<DefaultAtom, String>>, name: &'b str) -> Option<&'a String> {
    if let Some(map) = map {
        map.get(&DefaultAtom::from(name))
    } else {
        None
    }
}

fn get_prp_def<'a>(map: &'a Option<HashMap<DefaultAtom, String>>, name: &str, default: &'a str) -> &'a str {
    if let Some(map) = map {
        if let Some(value) = map.get(&DefaultAtom::from(name)) {
            value.as_ref()
        } else {
            default
        }
    } else {
        default
    }
}

