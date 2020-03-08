//! Implements reading and writing of ODS Files.
//!
//! Warning ahead: This does'nt cover the full specification, just a
//! useable subset to read + modify + write back an ODS file.
//!
//! ```
//! use spreadsheet_ods::{WorkBook, Sheet, Style, Family, ValueFormat, ValueType, FormatPart};
//! use chrono::NaiveDate;
//! use spreadsheet_ods::format;
//!
//! let mut wb = spreadsheet_ods::ods::read_ods("example.ods").unwrap();
//!
//! let mut sheet = wb.sheet_mut(0);
//! sheet.set_value(0, 0, 21.4f32);
//! sheet.set_value(0, 1, "foo");
//! sheet.set_styled_value(0, 2, NaiveDate::from_ymd(2020, 03, 01), "nice_date_style");
//! sheet.set_formula(0, 3, format!("of:={}+1", Sheet::fcellref(0,0)));
//!
//! let nice_date_format = format::create_date_mdy_format("nice_date_format");
//! wb.add_format(nice_date_format);
//!
//! let nice_date_style = Style::with_name(Family::TableCell, "nice_date_style", "nice_date_format");
//! wb.add_style(nice_date_style);
//!
//! spreadsheet_ods::ods::write_ods(&wb, "tryout.ods");
//!
//! ```
//!
//! When saving all the extra content is copied from the original file,
//! except for content.xml which is rewritten.
//!
//! For context.xml the following information is read and written:
//! * fonts
//! * styles
//! * table-data and structure
//!
//! The following things are ignored for now
//! * conditional formats
//! * charts
//! * ...
//!
//!

use std::collections::{BTreeMap, HashMap};
use std::collections::btree_map::Range;
use std::fmt;
use std::path::PathBuf;
use chrono::{Duration, NaiveDateTime, NaiveDate};
use rust_decimal::Decimal;
use rust_decimal::prelude::*;
use string_cache::DefaultAtom;

pub mod ods;
pub mod style;
pub mod defaultstyles;
pub mod format;

/// Book is the main structure for the Spreadsheet.
#[derive(Clone, Default)]
pub struct WorkBook {
    /// The data.
    sheets: Vec<Sheet>,

    //// FontDecl hold the style:font-face elements
    fonts: BTreeMap<String, FontDecl>,

    /// Styles hold the style:style elements.
    styles: BTreeMap<String, Style>,

    /// Value-styles are actual formatting instructions
    /// for various datatypes.
    /// Represents the various number:xxx-style elements.
    formats: BTreeMap<String, ValueFormat>,

    /// Default-styles per Type.
    /// This is only used when writing the ods file.
    def_styles: Option<HashMap<ValueType, String>>,

    /// Original file if this book was read from one.
    /// This is used for writing to copy all additional
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
            fonts: BTreeMap::new(),
            styles: BTreeMap::new(),
            formats: BTreeMap::new(),
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

    /// Creates a new sheet at the end and returns it.
    pub fn new_sheet(&mut self) -> &mut Sheet {
        self.sheets.push(Sheet::new());
        self.sheets.last_mut().unwrap()
    }

    /// Replaces the existing sheet.
    pub fn insert_sheet(&mut self, i: usize, sheet: Sheet) {
        self.sheets.insert(i, sheet);
    }

    /// Adds a sheet.
    pub fn push_sheet(&mut self, sheet: Sheet) {
        self.sheets.push(sheet);
    }

    /// Removes a sheet.
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

    // Finds a ValueStyle starting with the stylename attached to a cell.
    pub fn find_value_style(&self, style_name: &str) -> Option<&ValueFormat> {
        if let Some(style) = self.styles.get(style_name) {
            if let Some(value_style_name) = style.value_style() {
                if let Some(value_style) = self.formats.get(value_style_name) {
                    return Some(&value_style);
                }
            }
        }

        None
    }

    /// Adds a style.
    pub fn add_font(&mut self, font: FontDecl) {
        self.fonts.insert(font.name.to_string(), font);
    }

    pub fn remove_font(&mut self, name: &str) {
        self.fonts.remove(name);
    }

    pub fn font(&self, name: &str) -> Option<&FontDecl> {
        self.fonts.get(name)
    }

    pub fn font_mut(&mut self, name: &str) -> Option<&mut FontDecl> {
        self.fonts.get_mut(name)
    }

    /// Adds a style.
    pub fn add_style(&mut self, style: Style) { self.styles.insert(style.name.to_string(), style); }

    pub fn remove_style(&mut self, name: &str) { self.styles.remove(name); }

    pub fn style(&self, name: &str) -> Option<&Style> { self.styles.get(name) }

    pub fn style_mut(&mut self, name: &str) -> Option<&mut Style> {
        self.styles.get_mut(name)
    }

    /// Adds a value format.
    pub fn add_format(&mut self, vstyle: ValueFormat) {
        self.formats.insert(vstyle.name.to_string(), vstyle);
    }

    pub fn remove_format(&mut self, name: &str) {
        self.formats.remove(name);
    }

    pub fn format(&self, name: &str) -> Option<&ValueFormat> {
        self.formats.get(name)
    }

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

    data: BTreeMap<(usize, usize), SCell>,

    col_style: Option<BTreeMap<usize, String>>,
    col_cell_style: Option<BTreeMap<usize, String>>,
    row_style: Option<BTreeMap<usize, String>>,
}

impl fmt::Debug for Sheet {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(f, "{:?} {:?}", self.name, self.style)?;
        for (k, v) in self.data.iter() {
            writeln!(f, "{:?} {:?}", k, v)?;
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

    /// Creates a cell-reference for use in formulas.
    pub fn fcellref(row: usize, col: usize) -> String {
        let mut col_str = String::new();

        let mut col2 = col;
        while col2 > 0 {
            let digit = (col % 26) as u8;
            let cc = (b'A' + digit) as char;
            col_str.insert(0, cc);
            col2 /= 26;
        }

        let mut cell = String::from("[.");
        cell.push_str(&col_str);
        cell.push_str(&(row + 1).to_string());
        cell.push_str("]");

        cell
    }

    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets the table-style
    pub fn set_style<V: Into<String>>(&mut self, style: V) {
        self.style = Some(style.into());
    }

    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    /// Column wide style.
    pub fn set_column_style<V: Into<String>>(&mut self, col: usize, style: V) {
        if self.col_style.is_none() {
            self.col_style = Some(BTreeMap::new());
        }
        if let Some(col_style) = &mut self.col_style {
            col_style.entry(col).or_insert_with(|| style.into());
        }
    }

    pub fn column_style(&self, col: usize) -> Option<&String> {
        if let Some(col_style) = &self.col_style {
            col_style.get(&col)
        } else {
            None
        }
    }

    /// Default cell style for this column.
    pub fn set_column_cell_style<V: Into<String>>(&mut self, col: usize, style: V) {
        if self.col_cell_style.is_none() {
            self.col_cell_style = Some(BTreeMap::new());
        }
        if let Some(col_cell_style) = &mut self.col_cell_style {
            col_cell_style.entry(col).or_insert_with(|| style.into());
        }
    }

    pub fn column_cell_style(&self, col: usize) -> Option<&String> {
        if let Some(col_cell_style) = &self.col_cell_style {
            col_cell_style.get(&col)
        } else {
            None
        }
    }

    /// Row style.
    pub fn set_row_style<V: Into<String>>(&mut self, row: usize, style: V) {
        if self.row_style.is_none() {
            self.row_style = Some(BTreeMap::new());
        }
        if let Some(row_style) = &mut self.row_style {
            row_style.entry(row).or_insert_with(|| style.into());
        }
    }

    pub fn row_style(&self, row: usize) -> Option<&String> {
        if let Some(row_style) = &self.row_style {
            row_style.get(&row)
        } else {
            None
        }
    }

    /// Returns a tuple of (max(row)+1, max(col)+1)
    pub fn used_grid_size(&self) -> (usize, usize) {
        let max = self.data.keys().fold((0, 0), |mut max, (r, c)| {
            max.0 = usize::max(max.0, *r);
            max.1 = usize::max(max.1, *c);
            max
        });

        (max.0 + 1, max.1 + 1)
    }

    /// Returns a row of data.
    pub fn row_cells(&self, row: usize) -> Range<(usize, usize), SCell> {
        self.data.range((row, 0)..(row + 1, 0))
    }

    /// Returns the cell if available.
    pub fn cell(&self, row: usize, col: usize) -> Option<&SCell> {
        self.data.get(&(row, col))
    }

    /// Returns a mutable reference to the cell.
    pub fn cell_mut(&mut self, row: usize, col: usize) -> Option<&mut SCell> {
        self.data.get_mut(&(row, col))
    }

    /// Creates an empty cell if the position is currently empty and returns
    /// a reference.
    pub fn create_cell(&mut self, row: usize, col: usize) -> &mut SCell {
        self.data.entry((row, col)).or_insert_with(SCell::new)
    }

    /// Adds a cell. Replaces an existing one.
    pub fn add_cell(&mut self, row: usize, col: usize, cell: SCell) -> Option<SCell> {
        self.data.insert((row, col), cell)
    }

    // Removes a value.
    pub fn remove_cell(&mut self, row: usize, col: usize) -> Option<SCell> {
        self.data.remove(&(row, col))
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_styled_value<V: Into<Value>, W: Into<String>>(&mut self, row: usize, col: usize, value: V, style: W) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.value = Some(value.into());
        cell.style = Some(style.into());
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_value<V: Into<Value>>(&mut self, row: usize, col: usize, value: V) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.value = Some(value.into());
    }

    /// Returns a value
    pub fn value(&self, row: usize, col: usize) -> Option<Value> {
        if let Some(cell) = self.data.get(&(row, col)) {
            cell.value.as_ref().cloned()
        } else {
            None
        }
    }

    /// Returns a value
    pub fn value_ref(&self, row: usize, col: usize) -> Option<&Value> {
        if let Some(cell) = self.data.get(&(row, col)) {
            cell.value.as_ref()
        } else {
            None
        }
    }

    /// Sets a formula for the specified cell. Creates a new cell if necessary.
    pub fn set_formula<V: Into<String>>(&mut self, row: usize, col: usize, formula: V) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.formula = Some(formula.into());
    }

    /// Returns a value
    pub fn formula(&self, row: usize, col: usize) -> Option<&String> {
        if let Some(c) = self.data.get(&(row, col)) {
            c.formula.as_ref()
        } else {
            None
        }
    }

    /// Sets the cell-style for the specified cell. Creates a new cell if necessary.
    pub fn set_cell_style<V: Into<String>>(&mut self, row: usize, col: usize, style: V) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.style = Some(style.into());
    }

    /// Returns a value
    pub fn cell_style(&self, row: usize, col: usize) -> Option<&String> {
        if let Some(c) = self.data.get(&(row, col)) {
            c.style.as_ref()
        } else {
            None
        }
    }
}

/// One Cell of the spreadsheet.
#[derive(Debug, Clone, Default)]
pub struct SCell {
    value: Option<Value>,
    /// Unparsed formula string.
    formula: Option<String>,
    /// Cell style name.
    style: Option<String>,
}

impl SCell {
    pub fn new() -> Self {
        SCell {
            value: None,
            formula: None,
            style: None,
        }
    }

    pub fn with_value<V: Into<Value>>(value: V) -> Self {
        SCell {
            value: Some(value.into()),
            formula: None,
            style: None,
        }
    }

    pub fn value(&self) -> Option<&Value> {
        self.value.as_ref()
    }

    pub fn set_value<V: Into<Value>>(&mut self, value: V) {
        self.value = Some(value.into());
    }

    pub fn formula(&self) -> Option<&String> {
        self.formula.as_ref()
    }

    pub fn set_formula<V: Into<String>>(&mut self, formula: V) {
        self.formula = Some(formula.into());
    }

    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    pub fn set_style<V: Into<String>>(&mut self, style: V) {
        self.style = Some(style.into());
    }
}

/// Datatypes
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ValueType {
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
            Value::Boolean(_) => ValueType::Boolean,
            Value::Number(_) => ValueType::Number,
            Value::Percentage(_) => ValueType::Percentage,
            Value::Currency(_, _) => ValueType::Currency,
            Value::Text(_) => ValueType::Text,
            Value::TimeDuration(_) => ValueType::TimeDuration,
            Value::DateTime(_) => ValueType::DateTime,
        }
    }

    pub fn currency(currency: &str, value: f64) -> Self {
        Value::Currency(currency.to_string(), value)
    }

    pub fn percentage(percent: f64) -> Self {
        Value::Percentage(percent)
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

impl From<NaiveDate> for Value {
    fn from(dt: NaiveDate) -> Self { Value::DateTime(dt.and_hms(0, 0, 0)) }
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
    origin: Origin,
    /// All other attributes.
    prp: Option<HashMap<DefaultAtom, String>>,
}

impl FontDecl {
    pub fn new() -> Self {
        FontDecl::new_origin(Origin::Content)
    }

    pub fn new_origin(origin: Origin) -> Self {
        Self {
            name: "".to_string(),
            origin,
            prp: None,
        }
    }

    pub fn with_name<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            origin: Origin::Content,
            prp: None,
        }
    }

    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    pub fn name(&self) -> &String {
        &self.name
    }

    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    pub fn prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.prp, name, default)
    }
}

/// Style data fashioned after the ODS spec. Might not be too different to XLSX, but I didn't
/// check. There's a lot more, but for now I will simply ignore it. Seems most of
/// the blanks are filled in with defaults when reading.
///
/// The actual property names are just simple strings for now, maybe I map the common ones to
/// consts.
#[derive(Debug, Clone, Default)]
pub struct Style {
    name: String,
    /// Nice String.
    display_name: Option<String>,
    /// From where did we get this style.
    origin: Origin,
    /// Applicability of this style.
    family: Family,
    /// Styles can cascade.
    parent: Option<String>,
    /// References the actual formatting instructions in the value-styles.
    value_style: Option<String>,
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
    pub fn new() -> Self {
        Style::new_origin(Origin::Content)
    }

    pub fn new_origin(origin: Origin) -> Self {
        Style {
            name: String::from(""),
            display_name: None,
            origin,
            family: Family::None,
            parent: None,
            value_style: None,
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    pub fn with_name<S: Into<String>>(family: Family, name: S, value_style: S) -> Self {
        Style {
            name: name.into(),
            display_name: None,
            origin: Origin::Content,
            family,
            parent: Some(String::from("Default")),
            value_style: Some(value_style.into()),
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    pub fn name(&self) -> &String {
        &self.name
    }

    pub fn set_display_name(&mut self, name: &str) {
        self.display_name = Some(name.to_string());
    }

    pub fn display_name(&self) -> Option<&String> {
        self.display_name.as_ref()
    }

    pub fn set_origin(&mut self, origin: Origin) {
        self.origin = origin;
    }

    pub fn origin(&self) -> &Origin {
        &self.origin
    }

    pub fn set_family(&mut self, family: Family) {
        self.family = family;
    }

    pub fn family(&self) -> &Family {
        &self.family
    }

    pub fn set_parent(&mut self, parent: &str) {
        self.parent = Some(parent.to_string());
    }

    pub fn parent(&self) -> Option<&String> {
        self.parent.as_ref()
    }

    pub fn set_value_style(&mut self, value_style: &str) {
        self.value_style = Some(value_style.to_string());
    }

    pub fn value_style(&self) -> Option<&String> {
        self.value_style.as_ref()
    }

    pub fn set_table_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_prp, name, value);
    }

    pub fn table_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_prp, name)
    }

    pub fn table_prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.table_prp, name, default)
    }

    pub fn set_table_col_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_col_prp, name, value);
    }

    pub fn table_col_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_col_prp, name)
    }

    pub fn table_col_prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.table_col_prp, name, default)
    }

    pub fn set_table_row_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_row_prp, name, value);
    }

    pub fn table_row_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_row_prp, name)
    }

    pub fn table_row_prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.table_row_prp, name, default)
    }

    pub fn set_table_cell_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_cell_prp, name, value);
    }

    pub fn table_cell_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_cell_prp, name)
    }

    pub fn table_cell_prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.table_cell_prp, name, default)
    }

    pub fn set_text_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.text_prp, name, value);
    }

    pub fn clear_text_prp(&mut self, name: &str) -> Option<String> {
        clear_prp(&mut self.text_prp, name)
    }

    pub fn text_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.text_prp, name)
    }

    pub fn text_prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.text_prp, name, default)
    }

    pub fn set_paragraph_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.paragraph_prp, name, value);
    }

    pub fn paragraph_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.paragraph_prp, name)
    }

    pub fn paragraph_prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.paragraph_prp, name, default)
    }
}

/// Origin of a style. Content.xml or Styles.xml.
#[derive(Debug, Clone, PartialEq)]
pub enum Origin {
    Content,
    Styles,
}

impl Default for Origin {
    fn default() -> Self {
        Origin::Content
    }
}

/// Applicability of this style.
#[derive(Debug, Clone, PartialEq)]
pub enum Family {
    Table,
    TableRow,
    TableColumn,
    TableCell,
    None,
}

impl Default for Family {
    fn default() -> Self {
        Family::None
    }
}

/// Actual textual formatting of values.
#[derive(Debug, Clone, Default)]
pub struct ValueFormat {
    name: String,
    v_type: ValueType,
    origin: Origin,
    prp: Option<HashMap<DefaultAtom, String>>,
    parts: Option<Vec<FormatPart>>,
}

impl ValueFormat {
    pub fn new() -> Self {
        ValueFormat::new_origin(Origin::Content)
    }

    pub fn new_origin(origin: Origin) -> Self {
        ValueFormat {
            name: String::from(""),
            v_type: ValueType::Text,
            origin,
            prp: None,
            parts: None,
        }
    }

    pub fn with_name<S: Into<String>>(name: S, value_type: ValueType) -> Self {
        ValueFormat {
            name: name.into(),
            v_type: value_type,
            origin: Origin::Content,
            prp: None,
            parts: None,
        }
    }

    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    pub fn name(&self) -> &String {
        &self.name
    }

    pub fn set_value_type(&mut self, value_type: ValueType) {
        self.v_type = value_type;
    }

    pub fn value_type(&self) -> &ValueType {
        &self.v_type
    }

    pub fn set_origin(&mut self, origin: Origin) {
        self.origin = origin;
    }

    pub fn origin(&self) -> &Origin {
        &self.origin
    }

    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    pub fn prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.prp, name, default)
    }

    pub fn push_part(&mut self, part: FormatPart) {
        if let Some(parts) = &mut self.parts {
            parts.push(part);
        } else {
            self.parts = Some(vec![part]);
        }
    }

    pub fn push_parts(&mut self, parts: Vec<FormatPart>) {
        for p in parts.into_iter() {
            self.push_part(p);
        }
    }

    pub fn parts(&self) -> Option<&Vec<FormatPart>> {
        self.parts.as_ref()
    }

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
            let h12 = parts.iter().any(|v| v.ftype == FormatType::AmPm);

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
pub enum FormatType {
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
    ftype: FormatType,
    prp: Option<HashMap<DefaultAtom, String>>,
    content: Option<String>,
}

impl FormatPart {
    pub fn new(ftype: FormatType) -> Self {
        FormatPart {
            ftype,
            prp: None,
            content: None,
        }
    }

    pub fn new_content(ftype: FormatType, content: &str) -> Self {
        FormatPart {
            ftype,
            prp: None,
            content: Some(content.to_string()),
        }
    }

    pub fn new_vec(ftype: FormatType, vec: Vec<(&str, String)>) -> Self {
        let mut part = FormatPart {
            ftype,
            prp: None,
            content: None,
        };
        part.set_prp_vec(vec);
        part
    }

    pub fn set_ftype(&mut self, p_type: FormatType) {
        self.ftype = p_type;
    }

    pub fn ftype(&self) -> &FormatType {
        &self.ftype
    }

    pub fn set_prp_vec(&mut self, vec: Vec<(&str, String)>) {
        set_prp_vec(&mut self.prp, vec);
    }

    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    pub fn prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.prp, name, default)
    }

    /// Sets a textual content for this part. This is only used
    /// for text and currency-symbol.
    pub fn set_content(&mut self, content: &str) {
        self.content = Some(content.to_string());
    }

    pub fn content(&self) -> Option<&String> {
        self.content.as_ref()
    }

    /// Tries to format the given boolean.
    /// If this part does'nt match does nothing
    pub fn format_boolean(&self, buf: &mut String, b: bool) {
        match self.ftype {
            FormatType::Boolean => {
                buf.push_str(if b { "true" } else { "false" });
            }
            FormatType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given float.
    /// If this part does'nt match does nothing
    pub fn format_float(&self, buf: &mut String, f: f64) {
        match self.ftype {
            FormatType::Number => {
                let dec = self.prp_def("number:decimal-places", "0").parse::<usize>();
                if let Ok(dec) = dec {
                    buf.push_str(&format!("{:.*}", dec, f));
                }
            }
            FormatType::Scientific => {
                buf.push_str(&format!("{:e}", f));
            }
            FormatType::CurrencySymbol => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            FormatType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given float.
    /// If this part does'nt match does nothing
    pub fn format_str(&self, buf: &mut String, s: &str) {
        match self.ftype {
            FormatType::TextContent => {
                buf.push_str(s);
            }
            FormatType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given DateTime.
    /// Uses chrono::strftime for the implementation.
    /// If this part does'nt match does nothing
    #[allow(clippy::collapsible_if)]
    pub fn format_datetime(&self, buf: &mut String, d: &NaiveDateTime, h12: bool) {
        match self.ftype {
            FormatType::Day => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%d").to_string());
                } else {
                    buf.push_str(&d.format("%-d").to_string());
                }
            }
            FormatType::Month => {
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
            FormatType::Year => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%Y").to_string());
                } else {
                    buf.push_str(&d.format("%y").to_string());
                }
            }
            FormatType::DayOfWeek => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%A").to_string());
                } else {
                    buf.push_str(&d.format("%a").to_string());
                }
            }
            FormatType::WeekOfYear => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%W").to_string());
                } else {
                    buf.push_str(&d.format("%-W").to_string());
                }
            }
            FormatType::Hours => {
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
            FormatType::Minutes => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%M").to_string());
                } else {
                    buf.push_str(&d.format("%-M").to_string());
                }
            }
            FormatType::Seconds => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%S").to_string());
                } else {
                    buf.push_str(&d.format("%-S").to_string());
                }
            }
            FormatType::AmPm => {
                buf.push_str(&d.format("%p").to_string());
            }
            FormatType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given Duration.
    /// If this part does'nt match does nothing
    pub fn format_time_duration(&self, buf: &mut String, d: &Duration) {
        match self.ftype {
            FormatType::Hours => {
                buf.push_str(&d.num_hours().to_string());
            }
            FormatType::Minutes => {
                buf.push_str(&(d.num_minutes() % 60).to_string());
            }
            FormatType::Seconds => {
                buf.push_str(&(d.num_seconds() % 60).to_string());
            }
            FormatType::Text => {
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

