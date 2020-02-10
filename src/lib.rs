use std::collections::{BTreeMap, HashMap};
use std::fmt;
use std::path::PathBuf;

use chrono::{Duration, NaiveDateTime};

pub use crate::ods::OdsError;
pub use crate::ods::read_ods;
pub use crate::ods::write_ods;

pub mod ods;

/// Book is the main structure for the Spreadsheet.
pub struct WorkBook {
    /// The data.
    sheets: Vec<Sheet>,

    /// Styles hold the style:style elements.
    styles: BTreeMap<String, Style>,

    /// Value-styles are actual formatting instructions
    /// for various datatypes.
    /// Represents the various number:xxx-style elements.
    value_styles: BTreeMap<String, ValueStyle>,

    /// Default-styles per Type.
    /// This is only used when writing the ods file.
    def_styles: Option<HashMap<ValueType, String>>,

    /// Original file if this book was read from one
    /// This is used for writing to copy all additional
    /// files except content.xml
    file: Option<PathBuf>,
}

impl fmt::Debug for WorkBook {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        for s in self.sheets.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.styles.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.value_styles.iter() {
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
            styles: BTreeMap::new(),
            value_styles: BTreeMap::new(),
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
    pub fn find_value_style(&self, style_name: &str) -> Option<&ValueStyle> {
        if let Some(style) = self.styles.get(style_name) {
            if let Some(value_style_name) = style.value_style() {
                if let Some(value_style) = self.value_styles.get(value_style_name) {
                    return Some(&value_style);
                }
            }
        }

        None
    }

    /// Adds a style.
    pub fn add_style(&mut self, style: Style) {
        self.styles.insert(style.name.to_string(), style);
    }

    pub fn remove_style(&mut self, name: &str) {
        self.styles.remove(name);
    }

    pub fn style(&self, name: &str) -> Option<&Style> {
        self.styles.get(name)
    }

    pub fn style_mut(&mut self, name: &str) -> Option<&mut Style> {
        self.styles.get_mut(name)
    }

    /// Adds a value format.
    pub fn add_value_style(&mut self, vstyle: ValueStyle) {
        self.value_styles.insert(vstyle.name.to_string(), vstyle);
    }

    pub fn remove_value_style(&mut self, name: &str) {
        self.value_styles.remove(name);
    }

    pub fn value_style(&self, name: &str) -> Option<&ValueStyle> {
        self.value_styles.get(name)
    }

    pub fn value_style_mut(&mut self, name: &str) -> Option<&mut ValueStyle> {
        self.value_styles.get_mut(name)
    }
}

/// Adds default-styles for all basic ValueTypes. These are also set as default
/// styles for the respective types. By calling this function for a new workbook,
/// the basic formatting is done.
///
/// Beware
/// There is no i18n yet, so currency is set to euro for now.
/// And dates are european DMY style.
///
pub fn create_default_styles(book: &mut WorkBook) {
    let mut s = ValueStyle::with_name(Origin::Content, "BOOL1", ValueType::Boolean);
    s.push_part(Part::new(PartType::Boolean));
    book.add_value_style(s);

    let mut s = ValueStyle::with_name(Origin::Content, "NUM1", ValueType::Number);
    s.push_part(Part::def_number(2, false));
    book.add_value_style(s);

    let mut s = ValueStyle::with_name(Origin::Content, "PERCENT1", ValueType::Percentage);
    s.push_parts(Part::def_percentage(2));
    book.add_value_style(s);

    let mut s = ValueStyle::with_name(Origin::Content, "CURRENCY1", ValueType::Currency);
    s.push_parts(Part::def_euro());
    book.add_value_style(s);

    let mut s = ValueStyle::with_name(Origin::Content, "DATE1", ValueType::DateTime);
    s.push_parts(Part::def_date());
    book.add_value_style(s);

    let mut s = ValueStyle::with_name(Origin::Content, "DATETIME1", ValueType::DateTime);
    s.push_parts(Part::def_datetime());
    book.add_value_style(s);

    let mut s = ValueStyle::with_name(Origin::Content, "TIME1", ValueType::TimeDuration);
    s.push_parts(Part::def_time());
    book.add_value_style(s);

    let s = Style::with_name(Origin::Content, Family::TableCell, "DEFAULT-BOOL", "BOOLEAN1");
    book.add_style(s);

    let s = Style::with_name(Origin::Content, Family::TableCell, "DEFAULT-NUM", "NUM1");
    book.add_style(s);

    let s = Style::with_name(Origin::Content, Family::TableCell, "DEFAULT-PERCENT", "PERCENT1");
    book.add_style(s);

    let s = Style::with_name(Origin::Content, Family::TableCell, "DEFAULT-CURRENCY", "CURRENCY1");
    book.add_style(s);

    let s = Style::with_name(Origin::Content, Family::TableCell, "DEFAULT-DATE", "DATE1");
    book.add_style(s);

    let s = Style::with_name(Origin::Content, Family::TableCell, "DEFAULT-TIME", "TIME1");
    book.add_style(s);

    book.add_def_style(ValueType::Boolean, "DEFAULT-BOOL");
    book.add_def_style(ValueType::Number, "DEFAULT-NUM");
    book.add_def_style(ValueType::Percentage, "DEFAULT-PERCENT");
    book.add_def_style(ValueType::Currency, "DEFAULT-CURRENCY");
    book.add_def_style(ValueType::DateTime, "DEFAULT-DATE");
    book.add_def_style(ValueType::TimeDuration, "DEFAULT-TIME");
}

/// One sheet of the spreadsheet.
///
/// Contains the data and the style-references. The can also be
/// styles on the whole sheet, columns and rows. The more complicated
/// grouping tags are not covered.
pub struct Sheet {
    name: String,
    style: Option<String>,

    data: BTreeMap<(usize, usize), SCell>,
    columns: BTreeMap<usize, SColumn>,
    rows: BTreeMap<usize, SRow>,
}

impl fmt::Debug for Sheet {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(f, "{:?} {:?}", self.name, self.style)?;
        for (k, v) in self.data.iter() {
            writeln!(f, "{:?} {:?}", k, v)?;
        }
        for (k, v) in self.columns.iter() {
            writeln!(f, "{:?} {:?}", k, v)?;
        }
        for (k, v) in self.rows.iter() {
            writeln!(f, "{:?} {:?}", k, v)?;
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
            columns: BTreeMap::new(),
            rows: BTreeMap::new(),
        }
    }

    // New, empty, but with a name.
    pub fn with_name<V: Into<String>>(name: V) -> Self {
        Sheet {
            name: name.into(),
            data: BTreeMap::new(),
            style: None,
            columns: BTreeMap::new(),
            rows: BTreeMap::new(),
        }
    }

    /// Creates a cell-reference for use in formulas.
    pub fn colrow(row: usize, col: usize) -> String {
        let mut col_str = String::new();

        let mut col2 = col;
        while col2 > 0 {
            let digit = (col % 26) as u8;
            let cc = ('A' as u8 + digit) as char;
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
        let v = self.columns.entry(col).or_insert(SColumn::new());
        v.style = Some(style.into());
    }

    pub fn column_style(&self, col: usize) -> Option<&String> {
        if let Some(column) = self.columns.get(&col) {
            column.style.as_ref()
        } else {
            None
        }
    }

    /// Default cell style for this column.
    pub fn set_column_cell_style<V: Into<String>>(&mut self, col: usize, style: V) {
        let v = self.columns.entry(col).or_insert(SColumn::new());
        v.def_cell_style = Some(String::from(style.into()));
    }

    pub fn column_cell_style(&self, col: usize) -> Option<&String> {
        if let Some(column) = self.columns.get(&col) {
            column.style.as_ref()
        } else {
            None
        }
    }

    /// Row style.
    pub fn set_row_style<V: Into<String>>(&mut self, row: usize, style: V) {
        let v = self.rows.entry(row).or_insert(SRow::new());
        v.style = Some(style.into());
    }

    pub fn row_style(&self, row: usize) -> Option<&String> {
        if let Some(row) = self.rows.get(&row) {
            row.style.as_ref()
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
        self.data.entry((row, col)).or_insert(SCell::new())
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
    pub fn set_value<V: Into<Value>>(&mut self, row: usize, col: usize, value: V) {
        let mut cell = self.data.entry((row, col)).or_insert(SCell::new());
        cell.value = Some(value.into());
    }

    /// Returns a value
    pub fn value(&self, row: usize, col: usize) -> Option<Value> {
        if let Some(cell) = self.data.get(&(row, col)) {
            cell.value.as_ref().map(|v| v.clone())
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
        let mut cell = self.data.entry((row, col)).or_insert(SCell::new());
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
        let mut cell = self.data.entry((row, col)).or_insert(SCell::new());
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

/// Row wide data.
#[derive(Clone, Debug)]
struct SRow {
    style: Option<String>,
}

impl SRow {
    pub fn new() -> Self {
        SRow {
            style: None,
        }
    }
}

/// Column wide data.
#[derive(Clone, Debug)]
struct SColumn {
    /// The table-column style itself.
    style: Option<String>,
    /// Default cell-style for this column.
    def_cell_style: Option<String>,
}

impl SColumn {
    pub fn new() -> Self {
        SColumn {
            style: None,
            def_cell_style: None,
        }
    }
}

/// One Cell of the spreadsheet.
#[derive(Debug)]
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

/// Content-Values
#[derive(Debug, Clone)]
pub enum Value {
    Boolean(bool),
    DateTime(NaiveDateTime),
    TimeDuration(Duration),
    Number(f64),
    Currency(String, f64),
    Percentage(f64),
    Text(String),
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

impl From<Duration> for Value {
    fn from(d: Duration) -> Self {
        Value::TimeDuration(d)
    }
}

/// Style data fashioned after the ODS spec. Might not be too different to XLSX, but I didn't
/// check. There's a lot more, but for now I will simply ignore it. Seems most of
/// the blanks are filled in with defaults when reading.
///
/// The actual property names are just simple strings for now, maybe I map the common ones to
/// consts.
#[derive(Debug)]
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

    table_prp: Option<HashMap<String, String>>,
    table_col_prp: Option<HashMap<String, String>>,
    table_row_prp: Option<HashMap<String, String>>,
    table_cell_prp: Option<HashMap<String, String>>,
    paragraph_prp: Option<HashMap<String, String>>,
    text_prp: Option<HashMap<String, String>>,
}

impl Style {
    pub fn new(origin: Origin) -> Self {
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

    pub fn with_name(origin: Origin, family: Family, name: &str, value_style: &str) -> Self {
        Style {
            name: name.to_string(),
            display_name: None,
            origin,
            family,
            parent: Some(String::from("Default")),
            value_style: Some(value_style.to_string()),

            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_string();
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
#[derive(Debug, PartialEq)]
pub enum Origin {
    Content,
    Styles,
    None,
}

/// Applicability of this style.
#[derive(Debug)]
pub enum Family {
    Table,
    TableRow,
    TableColumn,
    TableCell,
    None,
}

/// Datatypes
#[derive(Debug, PartialEq, Eq, Hash)]
pub enum ValueType {
    Boolean,
    Number,
    Percentage,
    Currency,
    Text,
    DateTime,
    TimeDuration,
}

/// Actual textual formatting of values.
#[derive(Debug)]
pub struct ValueStyle {
    name: String,
    v_type: ValueType,
    origin: Origin,

    prp: Option<HashMap<String, String>>,

    parts: Option<Vec<Part>>,
}

impl ValueStyle {
    pub fn new(origin: Origin) -> Self {
        ValueStyle {
            name: String::from(""),
            v_type: ValueType::Text,
            origin,
            prp: None,
            parts: None,
        }
    }

    pub fn with_name(origin: Origin, name: &str, value_type: ValueType) -> Self {
        ValueStyle {
            name: name.to_string(),
            v_type: value_type,
            origin,
            prp: None,
            parts: None,
        }
    }

    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_string();
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

    pub fn push_part(&mut self, part: Part) {
        if let Some(parts) = &mut self.parts {
            parts.push(part);
        } else {
            self.parts = Some(vec![part]);
        }
    }

    pub fn push_parts(&mut self, parts: Vec<Part>) {
        for p in parts.into_iter() {
            self.push_part(p);
        }
    }

    pub fn parts(&self) -> Option<&Vec<Part>> {
        self.parts.as_ref()
    }

    pub fn parts_mut(&mut self) -> &mut Vec<Part> {
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
            let h12 = parts.iter().find(|v| v.p_type == PartType::AmPm).is_some();

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

/// The particles of a value->string format.
#[derive(Debug, PartialEq)]
pub enum PartType {
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

/// The particles of a value->string format.
#[derive(Debug)]
pub struct Part {
    p_type: PartType,
    prp: Option<HashMap<String, String>>,
    content: Option<String>,
}

impl Part {
    pub fn new(p_type: PartType) -> Self {
        Part {
            p_type,
            prp: None,
            content: None,
        }
    }

    pub fn new_content(p_type: PartType, content: &str) -> Self {
        Part {
            p_type,
            prp: None,
            content: Some(content.to_string()),
        }
    }

    pub fn new_prp(p_type: PartType, prp: HashMap<String, String>) -> Self {
        Part {
            p_type,
            prp: Some(prp),
            content: None,
        }
    }

    pub fn new_vec(p_type: PartType, vec: Vec<(&str, String)>) -> Self {
        let mut part = Part {
            p_type,
            prp: None,
            content: None,
        };
        part.set_prp_vec(vec);
        part
    }

    pub fn set_parttype(&mut self, p_type: PartType) {
        self.p_type = p_type;
    }

    pub fn parttype(&self) -> &PartType {
        &self.p_type
    }

    pub fn set_prp_map(&mut self, map: HashMap<String, String>) {
        self.prp = Some(map);
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
        match self.p_type {
            PartType::Boolean => {
                buf.push_str(if b { "true" } else { "false" });
            }
            PartType::Text => {
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
        match self.p_type {
            PartType::Number => {
                let dec = self.prp_def("number:decimal-places", "0").parse::<usize>();
                if let Ok(dec) = dec {
                    buf.push_str(&format!("{:.*}", dec, f));
                }
            }
            PartType::Scientific => {
                buf.push_str(&format!("{:e}", f));
            }
            PartType::CurrencySymbol => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            PartType::Text => {
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
        match self.p_type {
            PartType::TextContent => {
                buf.push_str(s);
            }
            PartType::Text => {
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
    pub fn format_datetime(&self, buf: &mut String, d: &NaiveDateTime, h12: bool) {
        match self.p_type {
            PartType::Day => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%d").to_string());
                } else {
                    buf.push_str(&d.format("%-d").to_string());
                }
            }
            PartType::Month => {
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
            PartType::Year => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%Y").to_string());
                } else {
                    buf.push_str(&d.format("%y").to_string());
                }
            }
            PartType::DayOfWeek => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%A").to_string());
                } else {
                    buf.push_str(&d.format("%a").to_string());
                }
            }
            PartType::WeekOfYear => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%W").to_string());
                } else {
                    buf.push_str(&d.format("%-W").to_string());
                }
            }
            PartType::Hours => {
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
            PartType::Minutes => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%M").to_string());
                } else {
                    buf.push_str(&d.format("%-M").to_string());
                }
            }
            PartType::Seconds => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%S").to_string());
                } else {
                    buf.push_str(&d.format("%-S").to_string());
                }
            }
            PartType::AmPm => {
                buf.push_str(&d.format("%p").to_string());
            }
            PartType::Text => {
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
        match self.p_type {
            PartType::Hours => {
                buf.push_str(&d.num_hours().to_string());
            }
            PartType::Minutes => {
                buf.push_str(&(d.num_minutes() % 60).to_string());
            }
            PartType::Seconds => {
                buf.push_str(&(d.num_seconds() % 60).to_string());
            }
            PartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Creates a new number format.
    pub fn def_number(decimal: u8, grouping: bool) -> Self {
        let mut p = Part::new(PartType::Number);
        p.set_prp("number:min-integer-digits", 1.to_string());
        p.set_prp("number:decimal-places", decimal.to_string());
        p.set_prp("loext:min-decimal-places", 0.to_string());
        if grouping {
            p.set_prp("number:grouping", String::from("true"));
        }
        p
    }

    /// Creates a new number format with a fixed number of decimal places.
    pub fn def_number_fixed(decimal: u8, grouping: bool) -> Self {
        let mut p = Part::new(PartType::Number);
        p.set_prp("number:min-integer-digits", 1.to_string());
        p.set_prp("number:decimal-places", decimal.to_string());
        p.set_prp("loext:min-decimal-places", decimal.to_string());
        if grouping {
            p.set_prp("number:grouping", String::from("true"));
        }
        p
    }

    /// Creates a new percantage format.
    pub fn def_percentage(decimal: u8) -> Vec<Self> {
        let mut p = Part::new(PartType::Number);
        p.set_prp("number:min-integer-digits", 1.to_string());
        p.set_prp("number:decimal-places", decimal.to_string());
        p.set_prp("loext:min-decimal-places", decimal.to_string());

        let mut p2 = Part::new(PartType::Text);
        p2.set_content("&#160;%");

        vec![p, p2]
    }

    /// Creates a new currency format for EURO.
    pub fn def_euro() -> Vec<Self> {
        let mut p0 = Part::new(PartType::CurrencySymbol);
        p0.set_prp("number:language", String::from("de"));
        p0.set_prp("number:country", String::from("AT"));
        p0.set_content("€");

        let mut p1 = Part::new(PartType::Text);
        p1.set_content(" ");

        let mut p2 = Part::new(PartType::Number);
        p2.set_prp("number:min-integer-digits", 1.to_string());
        p2.set_prp("number:decimal-places", 2.to_string());
        p2.set_prp("loext:min-decimal-places", 2.to_string());
        p2.set_prp("number:grouping", String::from("true"));

        vec![p0, p1, p2]
    }

    /// Creates a new currency format for EURO with negative values in red.
    /// Needs the name of the positive format.
    pub fn def_euro_red(gte0_style: &str) -> Vec<Self> {
        let mut p0 = Part::new(PartType::StyleText);
        p0.set_prp("fo:color", String::from("#ff0000"));

        let mut p1 = Part::new(PartType::Text);
        p1.set_content("-");

        let mut p2 = Part::new(PartType::CurrencySymbol);
        p2.set_prp("number:language", String::from("de"));
        p2.set_prp("number:country", String::from("AT"));
        p2.set_content("€");

        let mut p3 = Part::new(PartType::Text);
        p3.set_content(" ");

        let mut p4 = Part::new(PartType::Number);
        p4.set_prp("number:min-integer-digits", 1.to_string());
        p4.set_prp("number:decimal-places", 2.to_string());
        p4.set_prp("loext:min-decimal-places", 2.to_string());
        p4.set_prp("number:grouping", String::from("true"));

        let mut p5 = Part::new(PartType::StyleMap);
        p5.set_prp("style:condition", String::from("value()&gt;=0"));
        p5.set_prp("style:apply-style-name", gte0_style.to_string());

        vec![p0, p1, p2, p3, p4, p5]
    }

    /// Creates a new date format D.M.Y
    pub fn def_date() -> Vec<Self> {
        vec![
            Part::new_vec(PartType::Day, vec![("number:style", String::from("long"))]),
            Part::new_content(PartType::Text, "."),
            Part::new_vec(PartType::Month, vec![("number:style", String::from("long"))]),
            Part::new_content(PartType::Text, "."),
            Part::new_vec(PartType::Year, vec![("number:style", String::from("long"))]),
        ]
    }

    /// Creates a datetime froamt D.M.Y H:M:S
    pub fn def_datetime() -> Vec<Self> {
        vec![
            Part::new_vec(PartType::Year, vec![("number:style", String::from("long"))]),
            Part::new_content(PartType::Text, " "),
            Part::new_vec(PartType::Month, vec![("number:style", String::from("long"))]),
            Part::new_content(PartType::Text, "."),
            Part::new_vec(PartType::Day, vec![("number:style", String::from("long"))]),
            Part::new_content(PartType::Text, " "),
            Part::new(PartType::Hours),
            Part::new_content(PartType::Text, ":"),
            Part::new(PartType::Minutes),
            Part::new_content(PartType::Text, ":"),
            Part::new(PartType::Seconds),
        ]
    }

    /// Creates a new time-Duration format H:M:S
    pub fn def_time() -> Vec<Self> {
        vec![
            Part::new(PartType::Hours),
            Part::new_content(PartType::Text, " "),
            Part::new(PartType::Minutes),
            Part::new_content(PartType::Text, " "),
            Part::new(PartType::Seconds),
        ]
    }
}

fn set_prp_vec<'a>(map: &mut Option<HashMap<String, String>>, vec: Vec<(&str, String)>) {
    if map.is_none() {
        map.replace(HashMap::new());
    }
    if let Some(map) = map {
        for (name, value) in vec {
            map.insert(name.to_string(), value);
        }
    }
}

fn set_prp<'a>(map: &mut Option<HashMap<String, String>>, name: &str, value: String) {
    if map.is_none() {
        map.replace(HashMap::new());
    }
    if let Some(map) = map {
        map.insert(name.to_string(), value);
    }
}

fn get_prp<'a, 'b>(map: &'a Option<HashMap<String, String>>, name: &'b str) -> Option<&'a String> {
    if let Some(map) = map {
        map.get(name)
    } else {
        None
    }
}

fn get_prp_def<'a>(map: &'a Option<HashMap<String, String>>, name: &str, default: &'a str) -> &'a str {
    if let Some(map) = map {
        if let Some(value) = map.get(name) {
            value.as_ref()
        } else {
            default
        }
    } else {
        default
    }
}

