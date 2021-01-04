//! Implements reading and writing of ODS Files.
//!
//! ```
//! use spreadsheet_ods::{WorkBook, Sheet, Value};
//! use chrono::NaiveDate;
//! use spreadsheet_ods::format;
//! use spreadsheet_ods::formula;
//! use spreadsheet_ods::{cm, mm};
//! use spreadsheet_ods::style::{CellStyle};
//! use color::Rgb;
//! use spreadsheet_ods::style::units::{TextRelief, Border, Length};
//!
//!
//! let path = std::path::Path::new("tests/example.ods");
//! let mut wb = if path.exists() {
//!     spreadsheet_ods::read_ods(path).unwrap()
//! } else {
//!     WorkBook::new()
//! };
//!
//!
//! if wb.num_sheets() == 0 {
//!     let mut sheet = Sheet::new();
//!     sheet.cell_mut(0, 0).set_value(true);
//!     wb.push_sheet(sheet);
//! }
//!
//! let sheet = wb.sheet(0);
//! let n = sheet.value(0,0).as_f64_or(0f64);
//! if let Value::Boolean(v) = sheet.value(1,1) {
//!     if *v {
//!         println!("was true");
//!     }
//! }
//!
//! if wb.num_sheets() == 1 {
//!     wb.push_sheet(Sheet::new());
//! }
//!
//! let date_format = format::create_date_dmy_format("date_format");
//! let date_format = wb.add_format(date_format);
//!
//! let mut date_style = CellStyle::new("nice_date_style", &date_format);
//! date_style.set_font_bold();
//! date_style.set_font_relief(TextRelief::Engraved);
//! date_style.set_border(mm!(0.2), Border::Dashed, Rgb::new(192, 72, 72));
//! let date_style_ref = wb.add_cellstyle(date_style);
//!
//! let mut sheet = wb.sheet_mut(1);
//! sheet.set_value(0, 0, 21.4f32);
//! sheet.set_value(0, 1, "foo");
//! sheet.set_styled_value(0, 2, NaiveDate::from_ymd(2020, 03, 01), &date_style_ref);
//! sheet.set_formula(0, 3, format!("of:={}+1", formula::fcellref(0,0)));
//!
//! let mut sheet = Sheet::new_with_name("sample");
//! sheet.set_value(5,5, "sample");
//! wb.push_sheet(sheet);
//!

//!
//! spreadsheet_ods::write_ods(&wb, "test_out/tryout.ods");
//!
//! ```
//! This does not cover the entire ODS spec.
//!
//! What is supported:
//! * Spread-sheets
//!   * Handles all datatypes
//!     * Uses time::Duration
//!     * Uses chrono::NaiveDate and NaiveDateTime
//!     * Supports rust_decimal::Decimal
//!   * Column/Row/Cell styles
//!   * Formulas
//!     * Only as strings, but support functions for cell/range references.
//!   * Row/Column spans
//!   * Header rows/columns, print ranges
//!   * Formatted text as xml text.
//!
//! * Formulas
//!   * Only as strings.
//!   * Utilities for cell/range references.
//!
//! * Styles
//!   * Default styles per data type.
//!   * Preserves all style attributes.
//!   * Table, row, column, cell, paragraph and text styles.
//!   * Stylemaps (basic support)
//!   * Support for *setting* most style attributes.
//!
//! * Value formatting
//!   * The whole set is available.
//!   * Utility functions for common formats.
//!
//! * Fonts
//!   * Preserves all font attributes.
//!   * Basic support for setting this stuff.
//!
//! * Page layouts
//!   * Style attributes
//!   * Header/footer content as XML text.
//!
//! * Cell/range references
//!   * Parsing and formatting
//!
//! What is not supported:
//! * Spreadsheets
//!   * Row and column grouping
//! * ...
//!
//! There are a number of features that are not parsed completely,
//! but which are stored as a XML structure. This might work as long as
//! these features don't refer to data that is no longer valid after
//! some modification. But they are written back to the ods.
//!
//! Anyway those are:
//! * tracked-changes
//! * variable-decls
//! * sequence-decls
//! * user-field-decls
//! * dde-connection-decls
//! * calculation-settings
//! * content-validations
//! * label-ranges
//! * named-expressions
//! * database-ranges
//! * data-pilot-tables
//! * consolidation
//! * dde-links
//! * table:desc
//! * table-source
//! * dde-source
//! * scenario
//! * forms
//! * shapes
//! * calcext:conditional-formats
//!
//! When storing a previously read ODS file, all the contained files
//! are copied to the new file, except styles.xml and content.xml.
//! For a new ODS file mimetype, manifest, manifest.rdf, meta.xml
//! are filled with minimal defaults. There is no way to set these
//! for now.
//!

#![doc(html_root_url = "https://docs.rs/spreadsheet-ods/0.4.0")]

#[macro_use]
mod attr_macro;
#[macro_use]
mod unit_macro;
#[macro_use]
mod ref_macro;
mod attrmap2;
pub mod defaultstyles;
pub mod error;
pub mod format;
pub mod formula;
mod io;
pub mod refs;
pub mod style;
pub mod text;
pub mod xmltree;

pub use crate::error::OdsError;
pub use crate::format::{ValueFormat, ValueFormatRef};
pub use crate::io::{read_ods, write_ods};
pub use crate::refs::{CellRange, CellRef, ColRange, RowRange};
pub use crate::style::units::{Angle, Length};
pub use crate::style::{CellStyle, CellStyleRef};

use crate::style::{
    ColStyle, ColStyleRef, FontFaceDecl, GraphicStyle, GraphicStyleRef, MasterPage, MasterPageRef,
    PageStyle, PageStyleRef, ParagraphStyle, ParagraphStyleRef, RowStyle, RowStyleRef, TableStyle,
    TableStyleRef, TextStyle, TextStyleRef,
};
use crate::text::TextTag;
use crate::xmltree::XmlTag;
use chrono::{NaiveDate, NaiveDateTime};
#[cfg(feature = "use_decimal")]
use rust_decimal::prelude::*;
#[cfg(feature = "use_decimal")]
use rust_decimal::Decimal;
use std::collections::{BTreeMap, HashMap};
use std::fmt;
use std::fmt::{Display, Formatter};
use std::path::PathBuf;
use std::str::FromStr;
use time::Duration;

/// Cell index type for row/column indexes.
#[allow(non_camel_case_types)]
pub type ucell = u32;

/// Book is the main structure for the Spreadsheet.
#[derive(Clone, Default)]
pub struct WorkBook {
    /// The data.
    sheets: Vec<Sheet>,

    // ODS Version
    version: String,

    //// FontDecl hold the style:font-face elements
    fonts: HashMap<String, FontFaceDecl>,

    /// Styles hold the style:style elements.
    tablestyles: HashMap<String, TableStyle>,
    rowstyles: HashMap<String, RowStyle>,
    colstyles: HashMap<String, ColStyle>,
    cellstyles: HashMap<String, CellStyle>,
    paragraphstyles: HashMap<String, ParagraphStyle>,
    textstyles: HashMap<String, TextStyle>,
    graphicstyles: HashMap<String, GraphicStyle>,

    /// Value-styles are actual formatting instructions
    /// for various datatypes.
    /// Represents the various number:xxx-style elements.
    formats: HashMap<String, ValueFormat>,

    /// Default-styles per Type.
    /// This is only used when writing the ods file.
    def_styles: HashMap<ValueType, String>,

    /// Page-layout data.
    pagestyles: HashMap<String, PageStyle>,
    masterpages: HashMap<String, MasterPage>,

    /// Original file if this book was read from one.
    /// This is used when writing to copy all additional
    /// files except content.xml
    file: Option<PathBuf>,

    /// other stuff ...
    extra: Vec<XmlTag>,
}

impl fmt::Debug for WorkBook {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(f, "{:?}", self.version)?;
        for s in self.sheets.iter() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.fonts.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.tablestyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.rowstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.colstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.cellstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.paragraphstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.graphicstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats.values() {
            writeln!(f, "{:?}", s)?;
        }
        for (t, s) in &self.def_styles {
            writeln!(f, "{:?} -> {:?}", t, s)?;
        }
        for s in self.pagestyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.masterpages.values() {
            writeln!(f, "{:?}", s)?;
        }
        for xtr in &self.extra {
            writeln!(f, "extras {:?}", xtr)?;
        }
        writeln!(f, "{:?}", self.file)?;
        Ok(())
    }
}

impl WorkBook {
    pub fn new() -> Self {
        WorkBook {
            sheets: Default::default(),
            version: "1.3".to_string(),
            fonts: Default::default(),
            tablestyles: Default::default(),
            rowstyles: Default::default(),
            colstyles: Default::default(),
            cellstyles: Default::default(),
            paragraphstyles: Default::default(),
            textstyles: Default::default(),
            graphicstyles: Default::default(),
            formats: Default::default(),
            def_styles: Default::default(),
            pagestyles: Default::default(),
            masterpages: Default::default(),
            file: None,
            extra: vec![],
        }
    }

    /// Set ODS version.
    pub fn version(&self) -> &String {
        &self.version
    }

    /// ODS version.
    pub fn set_version(&mut self, version: String) {
        self.version = version;
    }

    /// Number of sheets.
    pub fn num_sheets(&self) -> usize {
        self.sheets.len()
    }

    /// Returns a certain sheet.
    ///
    /// Panics
    ///
    /// Panics if n does not exist.
    pub fn sheet(&self, n: usize) -> &Sheet {
        &self.sheets[n]
    }

    /// Returns a certain sheet.
    ///
    /// Panic
    ///
    /// Panics if n does not exist.
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
    pub fn add_def_style(&mut self, value_type: ValueType, style: &CellStyleRef) {
        self.def_styles.insert(value_type, style.to_string());
    }

    /// Returns the default style name.
    pub fn def_style(&self, value_type: ValueType) -> Option<&String> {
        self.def_styles.get(&value_type)
    }

    /// Finds a ValueFormat starting with the stylename attached to a cell.
    pub fn find_value_format(&self, style_name: &str) -> Option<&ValueFormat> {
        if let Some(style) = self.cellstyles.get(style_name) {
            if let Some(value_format_name) = style.value_format() {
                if let Some(value_format) = self.formats.get(value_format_name) {
                    return Some(&value_format);
                }
            }
        }

        None
    }

    /// Adds a font.
    pub fn add_font(&mut self, font: FontFaceDecl) {
        self.fonts.insert(font.name().to_string(), font);
    }

    /// Removes a font.
    pub fn remove_font(&mut self, name: &str) -> Option<FontFaceDecl> {
        self.fonts.remove(name)
    }

    /// Returns the FontDecl.
    pub fn font(&self, name: &str) -> Option<&FontFaceDecl> {
        self.fonts.get(name)
    }

    /// Returns a mutable FontDecl.
    pub fn font_mut(&mut self, name: &str) -> Option<&mut FontFaceDecl> {
        self.fonts.get_mut(name)
    }

    /// Adds a style.
    pub fn add_tablestyle(&mut self, style: TableStyle) -> TableStyleRef {
        let sref = style.style_ref();
        self.tablestyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_tablestyle(&mut self, name: &str) -> Option<TableStyle> {
        self.tablestyles.remove(name)
    }

    /// Returns the style.
    pub fn tablestyle(&self, name: &str) -> Option<&TableStyle> {
        self.tablestyles.get(name)
    }

    /// Returns the mutable style.
    pub fn tablestyle_mut(&mut self, name: &str) -> Option<&mut TableStyle> {
        self.tablestyles.get_mut(name)
    }

    /// Adds a style.
    pub fn add_rowstyle(&mut self, style: RowStyle) -> RowStyleRef {
        let sref = style.style_ref();
        self.rowstyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_rowstyle(&mut self, name: &str) -> Option<RowStyle> {
        self.rowstyles.remove(name)
    }

    /// Returns the style.
    pub fn rowstyle(&self, name: &str) -> Option<&RowStyle> {
        self.rowstyles.get(name)
    }

    /// Returns the mutable style.
    pub fn rowstyle_mut(&mut self, name: &str) -> Option<&mut RowStyle> {
        self.rowstyles.get_mut(name)
    }

    /// Adds a style.
    pub fn add_colstyle(&mut self, style: ColStyle) -> ColStyleRef {
        let sref = style.style_ref();
        self.colstyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_colstyle(&mut self, name: &str) -> Option<ColStyle> {
        self.colstyles.remove(name)
    }

    /// Returns the style.
    pub fn colstyle(&self, name: &str) -> Option<&ColStyle> {
        self.colstyles.get(name)
    }

    /// Returns the mutable style.
    pub fn colstyle_mut(&mut self, name: &str) -> Option<&mut ColStyle> {
        self.colstyles.get_mut(name)
    }

    /// Adds a style.
    pub fn add_cellstyle(&mut self, style: CellStyle) -> CellStyleRef {
        let sref = style.style_ref();
        self.cellstyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_cellstyle(&mut self, name: &str) -> Option<CellStyle> {
        self.cellstyles.remove(name)
    }

    /// Returns the style.
    pub fn cellstyle(&self, name: &str) -> Option<&CellStyle> {
        self.cellstyles.get(name)
    }

    /// Returns the mutable style.
    pub fn cellstyle_mut(&mut self, name: &str) -> Option<&mut CellStyle> {
        self.cellstyles.get_mut(name)
    }

    /// Adds a style.
    pub fn add_paragraphstyle(&mut self, style: ParagraphStyle) -> ParagraphStyleRef {
        let sref = style.style_ref();
        self.paragraphstyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_paragraphstyle(&mut self, name: &str) -> Option<ParagraphStyle> {
        self.paragraphstyles.remove(name)
    }

    /// Returns the style.
    pub fn paragraphstyle(&self, name: &str) -> Option<&ParagraphStyle> {
        self.paragraphstyles.get(name)
    }

    /// Returns the mutable style.
    pub fn paragraphstyle_mut(&mut self, name: &str) -> Option<&mut ParagraphStyle> {
        self.paragraphstyles.get_mut(name)
    }

    /// Adds a style.
    pub fn add_textstyle(&mut self, style: TextStyle) -> TextStyleRef {
        let sref = style.style_ref();
        self.textstyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_textstyle(&mut self, name: &str) -> Option<TextStyle> {
        self.textstyles.remove(name)
    }

    /// Returns the style.
    pub fn textstyle(&self, name: &str) -> Option<&TextStyle> {
        self.textstyles.get(name)
    }

    /// Returns the mutable style.
    pub fn textstyle_mut(&mut self, name: &str) -> Option<&mut TextStyle> {
        self.textstyles.get_mut(name)
    }

    /// Adds a style.
    pub fn add_graphicstyle(&mut self, style: GraphicStyle) -> GraphicStyleRef {
        let sref = style.style_ref();
        self.graphicstyles
            .insert(style.name().unwrap().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_graphicstyle(&mut self, name: &str) -> Option<GraphicStyle> {
        self.graphicstyles.remove(name)
    }

    /// Returns the style.
    pub fn graphicstyle(&self, name: &str) -> Option<&GraphicStyle> {
        self.graphicstyles.get(name)
    }

    /// Returns the mutable style.
    pub fn graphicstyle_mut(&mut self, name: &str) -> Option<&mut GraphicStyle> {
        self.graphicstyles.get_mut(name)
    }

    /// Adds a value format.
    pub fn add_format(&mut self, vstyle: ValueFormat) -> ValueFormatRef {
        let sref = vstyle.format_ref();
        self.formats.insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_format(&mut self, name: &str) -> Option<ValueFormat> {
        self.formats.remove(name)
    }

    /// Returns the format.
    pub fn format(&self, name: &str) -> Option<&ValueFormat> {
        self.formats.get(name)
    }

    /// Returns the mutable format.
    pub fn format_mut(&mut self, name: &str) -> Option<&mut ValueFormat> {
        self.formats.get_mut(name)
    }

    /// Adds a value PageStyle.
    pub fn add_pagestyle(&mut self, pstyle: PageStyle) -> PageStyleRef {
        let sref = pstyle.style_ref();
        self.pagestyles.insert(pstyle.name().to_string(), pstyle);
        sref
    }

    /// Removes the PageStyle.
    pub fn remove_pagestyle(&mut self, name: &str) -> Option<PageStyle> {
        self.pagestyles.remove(name)
    }

    /// Returns the PageStyle.
    pub fn pagestyle(&self, name: &str) -> Option<&PageStyle> {
        self.pagestyles.get(name)
    }

    /// Returns the mutable PageStyle.
    pub fn pagestyle_mut(&mut self, name: &str) -> Option<&mut PageStyle> {
        self.pagestyles.get_mut(name)
    }

    /// Adds a value MasterPage.
    pub fn add_masterpage(&mut self, mpage: MasterPage) -> MasterPageRef {
        let sref = mpage.masterpage_ref();
        self.masterpages.insert(mpage.name().to_string(), mpage);
        sref
    }

    /// Removes the MasterPage.
    pub fn remove_masterpage(&mut self, name: &str) -> Option<MasterPage> {
        self.masterpages.remove(name)
    }

    /// Returns the MasterPage.
    pub fn masterpage(&self, name: &str) -> Option<&MasterPage> {
        self.masterpages.get(name)
    }

    /// Returns the mutable MasterPage.
    pub fn masterpage_mut(&mut self, name: &str) -> Option<&mut MasterPage> {
        self.masterpages.get_mut(name)
    }
}

/// Visibility of a column or row.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum Visibility {
    Visible,
    Collapsed,
    Filtered,
}

impl FromStr for Visibility {
    type Err = OdsError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "visible" => Ok(Visibility::Visible),
            "filter" => Ok(Visibility::Filtered),
            "collapse" => Ok(Visibility::Collapsed),
            _ => Err(OdsError::Ods(format!(
                "Unknown value for table:visibility {}",
                s
            ))),
        }
    }
}

impl Default for Visibility {
    fn default() -> Self {
        Visibility::Visible
    }
}

impl Display for Visibility {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Visibility::Visible => write!(f, "visible"),
            Visibility::Collapsed => write!(f, "collapse"),
            Visibility::Filtered => write!(f, "filter"),
        }
    }
}

/// Row data
#[derive(Debug, Clone, Default)]
struct RowHeader {
    style: Option<String>,
    cellstyle: Option<String>,
    visible: Visibility,
}

impl RowHeader {
    pub fn new() -> Self {
        Self {
            style: None,
            cellstyle: None,
            visible: Default::default(),
        }
    }

    pub fn set_style(&mut self, style: &RowStyleRef) {
        self.style = Some(style.to_string());
    }

    pub fn clear_style(&mut self) {
        self.style = None;
    }

    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    pub fn set_cellstyle(&mut self, style: &CellStyleRef) {
        self.cellstyle = Some(style.to_string());
    }

    pub fn clear_cellstyle(&mut self) {
        self.cellstyle = None;
    }

    pub fn cellstyle(&self) -> Option<&String> {
        self.cellstyle.as_ref()
    }

    pub fn set_visible(&mut self, visible: Visibility) {
        self.visible = visible;
    }

    pub fn visible(&self) -> Visibility {
        self.visible
    }
}

/// Column data
#[derive(Debug, Clone, Default)]
struct ColHeader {
    style: Option<String>,
    cellstyle: Option<String>,
    visible: Visibility,
}

impl ColHeader {
    pub fn new() -> Self {
        Self {
            style: None,
            cellstyle: None,
            visible: Default::default(),
        }
    }

    pub fn set_style(&mut self, style: &ColStyleRef) {
        self.style = Some(style.to_string());
    }

    pub fn clear_style(&mut self) {
        self.style = None;
    }

    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    pub fn set_cellstyle(&mut self, style: &CellStyleRef) {
        self.cellstyle = Some(style.to_string());
    }

    pub fn clear_cellstyle(&mut self) {
        self.cellstyle = None;
    }

    pub fn cellstyle(&self) -> Option<&String> {
        self.cellstyle.as_ref()
    }

    pub fn set_visible(&mut self, visible: Visibility) {
        self.visible = visible;
    }

    pub fn visible(&self) -> Visibility {
        self.visible
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

    col_header: BTreeMap<ucell, ColHeader>,
    row_header: BTreeMap<ucell, RowHeader>,

    display: bool,
    print: bool,

    header_rows: Option<RowRange>,
    header_cols: Option<ColRange>,
    print_ranges: Option<Vec<CellRange>>,

    extra: Vec<XmlTag>,
}

impl fmt::Debug for Sheet {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(f, "name {:?} style {:?}", self.name, self.style)?;
        for (k, v) in self.data.iter() {
            writeln!(f, "  data {:?} {:?}", k, v)?;
        }
        for (k, v) in &self.col_header {
            writeln!(f, "{:?} {:?}", k, v)?;
        }
        for (k, v) in &self.row_header {
            writeln!(f, "{:?} {:?}", k, v)?;
        }
        if let Some(header_rows) = &self.header_rows {
            writeln!(f, "header rows {:?}", header_rows)?;
        }
        if let Some(header_cols) = &self.header_cols {
            writeln!(f, "header cols {:?}", header_cols)?;
        }
        for xtr in &self.extra {
            writeln!(f, "extras {:?}", xtr)?;
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
            col_header: Default::default(),
            style: None,
            header_rows: None,
            header_cols: None,
            print_ranges: None,
            extra: vec![],
            row_header: Default::default(),
            display: true,
            print: true,
        }
    }

    // New, empty, but with a name.
    pub fn new_with_name<S: Into<String>>(name: S) -> Self {
        Sheet {
            name: name.into(),
            data: BTreeMap::new(),
            col_header: Default::default(),
            style: None,
            header_rows: None,
            header_cols: None,
            print_ranges: None,
            extra: vec![],
            row_header: Default::default(),
            display: true,
            print: true,
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
    pub fn set_style(&mut self, style: &TableStyleRef) {
        self.style = Some(style.to_string());
    }

    /// Returns the table-style.
    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    /// Column style.
    pub fn set_colstyle(&mut self, col: ucell, style: &ColStyleRef) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_style(style);
    }

    /// Remove the style.
    pub fn clear_colstyle(&mut self, col: ucell) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .clear_style();
    }

    /// Returns the column style.
    pub fn colstyle(&self, col: ucell) -> Option<&String> {
        if let Some(col_header) = self.col_header.get(&col) {
            col_header.style()
        } else {
            None
        }
    }

    /// Default cell style for this column.
    pub fn set_col_cellstyle(&mut self, col: ucell, style: &CellStyleRef) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_cellstyle(style);
    }

    /// Remove the style.
    pub fn clear_col_cellstyle(&mut self, col: ucell) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .clear_cellstyle();
    }

    /// Returns the default cell style for this column.
    pub fn col_cellstyle(&self, col: ucell) -> Option<&String> {
        if let Some(col_header) = self.col_header.get(&col) {
            col_header.cellstyle()
        } else {
            None
        }
    }

    /// Visibility of the column
    pub fn set_col_visible(&mut self, col: ucell, visible: Visibility) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_visible(visible);
    }

    /// Returns the default cell style for this column.
    pub fn col_visible(&self, col: ucell) -> Visibility {
        if let Some(col_header) = self.col_header.get(&col) {
            col_header.visible()
        } else {
            Default::default()
        }
    }

    /// Creates a col style and sets the col width.
    pub fn set_col_width(&mut self, workbook: &mut WorkBook, col: ucell, width: Length) {
        let style_name = format!("co{}", col);

        let mut colstyle = if let Some(style) = workbook.remove_colstyle(&style_name) {
            style
        } else {
            ColStyle::new(&style_name)
        };
        colstyle.set_col_width(width);
        colstyle.set_use_optimal_col_width(false);
        let colstyle = workbook.add_colstyle(colstyle);

        self.set_colstyle(col, &colstyle);
    }

    /// Row style.
    pub fn set_rowstyle(&mut self, col: ucell, style: &RowStyleRef) {
        self.row_header
            .entry(col)
            .or_insert_with(RowHeader::new)
            .set_style(style);
    }

    /// Remove the style.
    pub fn clear_rowstyle(&mut self, col: ucell) {
        self.row_header
            .entry(col)
            .or_insert_with(RowHeader::new)
            .clear_style();
    }

    /// Returns the row style.
    pub fn rowstyle(&self, col: ucell) -> Option<&String> {
        if let Some(row_header) = self.row_header.get(&col) {
            row_header.style()
        } else {
            None
        }
    }

    /// Default cell style for this row.
    pub fn set_row_cellstyle(&mut self, col: ucell, style: &CellStyleRef) {
        self.row_header
            .entry(col)
            .or_insert_with(RowHeader::new)
            .set_cellstyle(style);
    }

    /// Remove the style.
    pub fn clear_row_cellstyle(&mut self, col: ucell) {
        self.row_header
            .entry(col)
            .or_insert_with(RowHeader::new)
            .clear_cellstyle();
    }

    /// Returns the default cell style for this row.
    pub fn row_cellstyle(&self, col: ucell) -> Option<&String> {
        if let Some(row_header) = self.row_header.get(&col) {
            row_header.cellstyle()
        } else {
            None
        }
    }

    /// Visibility of the row
    pub fn set_row_visible(&mut self, col: ucell, visible: Visibility) {
        self.row_header
            .entry(col)
            .or_insert_with(RowHeader::new)
            .set_visible(visible);
    }

    /// Returns the default cell style for this row.
    pub fn row_visible(&self, col: ucell) -> Visibility {
        if let Some(row_header) = self.row_header.get(&col) {
            row_header.visible()
        } else {
            Default::default()
        }
    }

    /// Creates a row-style and sets the row height.
    pub fn set_row_height(&mut self, workbook: &mut WorkBook, row: ucell, height: Length) {
        let style_name = format!("ro{}", row);

        let mut rowstyle = if let Some(style) = workbook.remove_rowstyle(&style_name) {
            style
        } else {
            RowStyle::new(&style_name)
        };
        rowstyle.set_row_height(height);
        rowstyle.set_use_optimal_row_height(false);
        let rowstyle = workbook.add_rowstyle(rowstyle);

        self.set_rowstyle(row, &rowstyle);
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

    /// Is the sheet displayed?
    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }

    /// Is the sheet displayed?
    pub fn display(&self) -> bool {
        self.display
    }

    /// Is the sheet printed?
    pub fn set_print(&mut self, print: bool) {
        self.print = print;
    }

    /// Is the sheet printed?
    pub fn print(&self) -> bool {
        self.print
    }

    /// Returns true if there is no SCell at the given position.
    pub fn is_empty(&self, row: ucell, col: ucell) -> bool {
        self.data.get(&(row, col)).is_none()
    }

    /// Returns the cell if available.
    pub fn cell(&self, row: ucell, col: ucell) -> Option<&SCell> {
        self.data.get(&(row, col))
    }

    /// Ensures that there is a SCell at the given position and returns
    /// a reference to it.
    pub fn cell_mut(&mut self, row: ucell, col: ucell) -> &mut SCell {
        self.data.entry((row, col)).or_insert_with(SCell::new)
    }

    /// Adds a cell. Replaces an existing one.
    pub fn add_cell(&mut self, row: ucell, col: ucell, cell: SCell) -> Option<SCell> {
        self.data.insert((row, col), cell)
    }

    /// Removes a cell.
    pub fn remove_cell(&mut self, row: ucell, col: ucell) -> Option<SCell> {
        self.data.remove(&(row, col))
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_styled_value<V: Into<Value>>(
        &mut self,
        row: ucell,
        col: ucell,
        value: V,
        style: &CellStyleRef,
    ) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.value = value.into();
        cell.style = Some(style.to_string());
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
    pub fn set_cellstyle(&mut self, row: ucell, col: ucell, style: &CellStyleRef) {
        let mut cell = self.data.entry((row, col)).or_insert_with(SCell::new);
        cell.style = Some(style.to_string());
    }

    /// Returns a value
    pub fn cellstyle(&self, row: ucell, col: ucell) -> Option<&String> {
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

    /// Rowspan of the cell.
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
        self.print_ranges.get_or_insert_with(Vec::new).push(range);
    }

    /// Remove print ranges.
    pub fn clear_print_ranges(&mut self) {
        self.print_ranges = None;
    }

    /// Return the print ranges.
    pub fn print_ranges(&self) -> Option<&Vec<CellRange>> {
        self.print_ranges.as_ref()
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
    pub fn set_style(&mut self, style: &CellStyleRef) {
        self.style = Some(style.to_string());
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ValueType {
    Empty,
    Boolean,
    Number,
    Percentage,
    Currency,
    Text,
    TextXml,
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
    TextXml(Box<TextTag>),
    DateTime(NaiveDateTime),
    TimeDuration(Duration),
}

impl Value {
    // Return the plan ValueType for this value.
    pub fn value_type(&self) -> ValueType {
        match self {
            Value::Empty => ValueType::Empty,
            Value::Boolean(_) => ValueType::Boolean,
            Value::Number(_) => ValueType::Number,
            Value::Percentage(_) => ValueType::Percentage,
            Value::Currency(_, _) => ValueType::Currency,
            Value::Text(_) => ValueType::Text,
            Value::TextXml(_) => ValueType::TextXml,
            Value::TimeDuration(_) => ValueType::TimeDuration,
            Value::DateTime(_) => ValueType::DateTime,
        }
    }

    /// Return the bool if the value is a Boolean. Default otherwise.
    pub fn as_bool_or(&self, d: bool) -> bool {
        match self {
            Value::Boolean(b) => *b,
            _ => d,
        }
    }

    /// Return the content as i32 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_i32_or(&self, d: i32) -> i32 {
        match self {
            Value::Number(n) => *n as i32,
            Value::Percentage(p) => *p as i32,
            Value::Currency(_, v) => *v as i32,
            _ => d,
        }
    }

    /// Return the content as i32 if the value is a number, percentage or
    /// currency.
    pub fn as_i32_opt(&self) -> Option<i32> {
        match self {
            Value::Number(n) => Some(*n as i32),
            Value::Percentage(p) => Some(*p as i32),
            Value::Currency(_, v) => Some(*v as i32),
            _ => None,
        }
    }

    /// Return the content as u32 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_u32_or(&self, d: u32) -> u32 {
        match self {
            Value::Number(n) => *n as u32,
            Value::Percentage(p) => *p as u32,
            Value::Currency(_, v) => *v as u32,
            _ => d,
        }
    }

    /// Return the content as u32 if the value is a number, percentage or
    /// currency.
    pub fn as_u32_opt(&self) -> Option<u32> {
        match self {
            Value::Number(n) => Some(*n as u32),
            Value::Percentage(p) => Some(*p as u32),
            Value::Currency(_, v) => Some(*v as u32),
            _ => None,
        }
    }

    /// Return the content as decimal if the value is a number, percentage or
    /// currency. Default otherwise.
    #[cfg(feature = "use_decimal")]
    pub fn as_decimal_or(&self, d: Decimal) -> Decimal {
        match self {
            Value::Number(n) => Decimal::from_f64(*n).unwrap(),
            Value::Currency(_, v) => Decimal::from_f64(*v).unwrap(),
            Value::Percentage(p) => Decimal::from_f64(*p).unwrap(),
            _ => d,
        }
    }

    /// Return the content as decimal if the value is a number, percentage or
    /// currency. Default otherwise.
    #[cfg(feature = "use_decimal")]
    pub fn as_decimal_opt(&self) -> Option<Decimal> {
        match self {
            Value::Number(n) => Some(Decimal::from_f64(*n).unwrap()),
            Value::Currency(_, v) => Some(Decimal::from_f64(*v).unwrap()),
            Value::Percentage(p) => Some(Decimal::from_f64(*p).unwrap()),
            _ => None,
        }
    }

    /// Return the content as f64 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_f64_or(&self, d: f64) -> f64 {
        match self {
            Value::Number(n) => *n,
            Value::Currency(_, v) => *v,
            Value::Percentage(p) => *p,
            _ => d,
        }
    }

    /// Return the content as f64 if the value is a number, percentage or
    /// currency.
    pub fn as_f64_opt(&self) -> Option<f64> {
        match self {
            Value::Number(n) => Some(*n),
            Value::Currency(_, v) => Some(*v),
            Value::Percentage(p) => Some(*p),
            _ => None,
        }
    }

    /// Return the content as str if the value is text.
    pub fn as_str_or<'a>(&'a self, d: &'a str) -> &'a str {
        match self {
            Value::Text(s) => s.as_ref(),
            _ => d,
        }
    }

    /// Return the content as str if the value is text.
    pub fn as_str_opt(&self) -> Option<&str> {
        match self {
            Value::Text(s) => Some(s.as_ref()),
            _ => None,
        }
    }

    /// Return the content as Duration if the value is a TimeDuration.
    /// Default otherwise.
    pub fn as_timeduration_or(&self, d: Duration) -> Duration {
        match self {
            Value::TimeDuration(td) => *td,
            _ => d,
        }
    }

    /// Return the content as Duration if the value is a TimeDuration.
    /// Default otherwise.
    pub fn as_timeduration_opt(&self) -> Option<Duration> {
        match self {
            Value::TimeDuration(td) => Some(*td),
            _ => None,
        }
    }

    /// Return the content as NaiveDateTime if the value is a DateTime.
    /// Default otherwise.
    pub fn as_datetime_or(&self, d: NaiveDateTime) -> NaiveDateTime {
        match self {
            Value::DateTime(dt) => *dt,
            _ => d,
        }
    }

    /// Return the content as an optional NaiveDateTime if the value is
    /// a DateTime.
    pub fn as_datetime_opt(&self) -> Option<NaiveDateTime> {
        match self {
            Value::DateTime(dt) => Some(*dt),
            _ => None,
        }
    }
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}

/// currency value
#[macro_export]
macro_rules! currency {
    ($c:expr, $v:expr) => {
        Value::Currency($c.to_string(), $v as f64)
    };
}

/// currency value
#[macro_export]
macro_rules! percent {
    ($v:expr) => {
        Value::Percentage($v)
    };
}

impl From<()> for Value {
    fn from(_: ()) -> Self {
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

impl From<TextTag> for Value {
    fn from(t: TextTag) -> Self {
        Value::TextXml(Box::new(t))
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

#[cfg(feature = "use_decimal")]
impl From<Decimal> for Value {
    fn from(f: Decimal) -> Self {
        Value::Number(f.to_f64().unwrap())
    }
}

#[cfg(feature = "use_decimal")]
impl From<Option<Decimal>> for Value {
    fn from(f: Option<Decimal>) -> Self {
        if let Some(f) = f {
            Value::Number(f.to_f64().unwrap())
        } else {
            Value::Empty
        }
    }
}

macro_rules! from_number {
    ($l:ty) => {
        impl From<$l> for Value {
            fn from(f: $l) -> Self {
                Value::Number(f as f64)
            }
        }

        impl From<Option<$l>> for Value {
            fn from(f: Option<$l>) -> Self {
                if let Some(f) = f {
                    Value::Number(f as f64)
                } else {
                    Value::Empty
                }
            }
        }
    };
}

from_number!(f64);
from_number!(f32);
from_number!(i64);
from_number!(i32);
from_number!(i16);
from_number!(i8);
from_number!(u64);
from_number!(u32);
from_number!(u16);
from_number!(u8);

impl From<bool> for Value {
    fn from(b: bool) -> Self {
        Value::Boolean(b)
    }
}

impl From<Option<bool>> for Value {
    fn from(b: Option<bool>) -> Self {
        if let Some(b) = b {
            Value::Boolean(b)
        } else {
            Value::Empty
        }
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
    fn from(dt: NaiveDate) -> Self {
        Value::DateTime(dt.and_hms(0, 0, 0))
    }
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

impl From<Option<Duration>> for Value {
    fn from(d: Option<Duration>) -> Self {
        if let Some(d) = d {
            Value::TimeDuration(d)
        } else {
            Value::Empty
        }
    }
}
