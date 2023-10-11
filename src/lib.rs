//! Implements reading and writing of ODS Files.
//!
//! ```
//! use spreadsheet_ods::{WorkBook, Sheet, Value};
//! use chrono::NaiveDate;
//! use spreadsheet_ods::format;
//! use spreadsheet_ods::formula;
//! use spreadsheet_ods::{cm, mm};
//! use spreadsheet_ods::style::{CellStyle};
//! use spreadsheet_ods::color::Rgb;
//! use icu_locid::locale;
//! use spreadsheet_ods::style::units::{TextRelief, Border, Length};
//!
//!
//! let path = std::path::Path::new("tests/example.ods");
//! let mut wb = if path.exists() {
//!     spreadsheet_ods::read_ods(path).unwrap()
//! } else {
//!     WorkBook::new(locale!("en_US"))
//! };
//!
//!
//! if wb.num_sheets() == 0 {
//!     let mut sheet = Sheet::new("simple");
//!     sheet.set_value(0, 0, true);
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
//!     wb.push_sheet(Sheet::new("two"));
//! }
//!
//! let date_format = format::create_date_dmy_format("date_format");
//! let date_format = wb.add_datetime_format(date_format);
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
//! sheet.set_styled_value(0, 2, NaiveDate::from_ymd_opt(2020, 03, 01), &date_style_ref);
//! sheet.set_formula(0, 3, format!("of:={}+1", formula::fcellref(0,0)));
//!
//! let mut sheet = Sheet::new("sample");
//! sheet.set_value(5,5, "sample");
//! wb.push_sheet(sheet);
//!
//!
//! spreadsheet_ods::write_ods(&mut wb, "test_out/tryout.ods");
//!
//! ```
//! This does not cover the entire ODS spec.
//!
//! What is supported:
//! * Spread-sheets
//!   * Handles all datatypes
//!     * Uses time::Duration
//!     * Uses chrono::NaiveDate and NaiveDateTime
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
//!   * Basic localization support.
//!
//! * Content validation
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
//! What might be problematic:
//! * The text content of each cell is not formatted according to the given ValueFormat,
//!   but instead is a simple to_string() of the data type. This data is not necessary
//!   to read the contents correctly. LibreOffice seems to ignore this completely
//!   and display everything correctly.
//!
//! Next on the TO-DO list:
//! * Calculation settings.
//! * Named expressions.
//!
//! There are a number of features that are not parsed to a structure,
//! but which are stored as a XML. This might work as long as
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
//! are copied to the new file. The files content.xml, styles.xml,
//! settings.xml and manifest.xml are written from the data.
//!

#![doc(html_root_url = "https://docs.rs/spreadsheet-ods")]
#![warn(absolute_paths_not_starting_with_crate)]
// NO #![warn(box_pointers)]
#![warn(elided_lifetimes_in_paths)]
#![warn(explicit_outlives_requirements)]
#![warn(keyword_idents)]
#![warn(macro_use_extern_crate)]
#![warn(meta_variable_misuse)]
#![warn(missing_abi)]
// NOT_ACCURATE #![warn(missing_copy_implementations)]
#![warn(missing_debug_implementations)]
#![warn(missing_docs)]
#![warn(non_ascii_idents)]
#![warn(noop_method_call)]
// NO #![warn(or_patterns_back_compat)]
#![warn(pointer_structural_match)]
#![warn(semicolon_in_expressions_from_macros)]
// NOT_ACCURATE #![warn(single_use_lifetimes)]
#![warn(trivial_casts)]
#![warn(trivial_numeric_casts)]
#![warn(unreachable_pub)]
// #![warn(unsafe_code)]
#![warn(unsafe_op_in_unsafe_fn)]
#![warn(unstable_features)]
// NO #![warn(unused_crate_dependencies)]
// NO #![warn(unused_extern_crates)]
#![warn(unused_import_braces)]
#![warn(unused_lifetimes)]
#![warn(unused_qualifications)]
// NO #![warn(unused_results)]
#![warn(variant_size_differences)]

pub use crate::error::OdsError;
pub use crate::format::{
    ValueFormatBoolean, ValueFormatCurrency, ValueFormatDateTime, ValueFormatNumber,
    ValueFormatPercentage, ValueFormatRef, ValueFormatText, ValueFormatTimeDuration,
};
pub use crate::io::read::{
    read_fods, read_fods_buf, read_fods_from, read_ods, read_ods_buf, read_ods_from, OdsOptions,
};
pub use crate::io::write::{
    write_fods, write_fods_buf, write_fods_to, write_ods, write_ods_buf,
    write_ods_buf_uncompressed, write_ods_to,
};
pub use crate::refs::{CellRange, CellRef, ColRange, RowRange};
pub use crate::style::units::{Angle, Length};
pub use crate::style::{CellStyle, CellStyleRef};
pub use color;

use crate::config::Config;
use crate::defaultstyles::{DefaultFormat, DefaultStyle};
use crate::draw::{Annotation, DrawFrame};
use crate::ds::detach::Detach;
use crate::ds::detach::Detached;
use crate::format::ValueFormatTrait;
use crate::io::read::default_settings;
use crate::manifest::Manifest;
use crate::metadata::Metadata;
use crate::style::{
    ColStyle, ColStyleRef, FontFaceDecl, GraphicStyle, GraphicStyleRef, MasterPage, MasterPageRef,
    PageStyle, PageStyleRef, ParagraphStyle, ParagraphStyleRef, RowStyle, RowStyleRef, RubyStyle,
    RubyStyleRef, TableStyle, TableStyleRef, TextStyle, TextStyleRef,
};
use crate::text::TextTag;
use crate::validation::{Validation, ValidationRef};
use crate::xlink::{XLinkActuate, XLinkType};
use crate::xmltree::{XmlContent, XmlTag};
use chrono::{Duration, NaiveTime};
use chrono::{NaiveDate, NaiveDateTime};
use icu_locid::{locale, Locale};
#[cfg(feature = "use_decimal")]
use rust_decimal::prelude::{FromPrimitive, ToPrimitive};
#[cfg(feature = "use_decimal")]
use rust_decimal::Decimal;
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::convert::TryFrom;
use std::fmt;
use std::fmt::{Display, Formatter};
use std::iter::FusedIterator;
use std::ops::RangeBounds;

#[macro_use]
mod macro_attr_draw;
#[macro_use]
mod macro_attr_style;
#[macro_use]
mod macro_attr_fo;
#[macro_use]
mod macro_attr_svg;
#[macro_use]
mod macro_attr_text;
#[macro_use]
mod macro_attr_number;
#[macro_use]
mod macro_attr_table;
#[macro_use]
mod macro_attr_xlink;
#[macro_use]
mod unit_macro;
#[macro_use]
mod format_macro;
#[macro_use]
mod style_macro;
#[macro_use]
mod text_macro;

mod attrmap2;
mod config;
mod ds;
mod io;
mod locale;

pub mod condition;
pub mod defaultstyles;
pub mod draw;
pub mod error;
pub mod format;
pub mod formula;
pub mod manifest;
pub mod metadata;
pub mod refs;
pub mod style;
pub mod text;
pub mod validation;
pub mod xlink;
pub mod xmltree;

// Use the IndexMap for debugging, makes diffing much easier.
// Otherwise the std::HashMap is good.
// pub(crate) type HashMap<K, V> = indexmap::IndexMap<K, V>;
// pub(crate) type HashMapIter<'a, K, V> = indexmap::map::Iter<'a, K, V>;
pub(crate) type HashMap<K, V> = std::collections::HashMap<K, V>;
pub(crate) type HashMapIter<'a, K, V> = std::collections::hash_map::Iter<'a, K, V>;

/// Book is the main structure for the Spreadsheet.
#[derive(Clone)]
pub struct WorkBook {
    /// The data.
    sheets: Vec<Detach<Sheet>>,

    /// ODS Version
    version: String,

    /// FontDecl hold the style:font-face elements
    fonts: HashMap<String, FontFaceDecl>,

    /// Auto-Styles. Maps the prefix to a number.
    autonum: HashMap<String, u32>,

    /// Scripts
    scripts: Vec<Script>,
    event_listener: HashMap<String, EventListener>,

    /// Styles hold the style:style elements.
    tablestyles: HashMap<String, TableStyle>,
    rowstyles: HashMap<String, RowStyle>,
    colstyles: HashMap<String, ColStyle>,
    cellstyles: HashMap<String, CellStyle>,
    paragraphstyles: HashMap<String, ParagraphStyle>,
    textstyles: HashMap<String, TextStyle>,
    rubystyles: HashMap<String, RubyStyle>,
    graphicstyles: HashMap<String, GraphicStyle>,

    /// Value-styles are actual formatting instructions for various datatypes.
    /// Represents the various number:xxx-style elements.
    formats_boolean: HashMap<String, ValueFormatBoolean>,
    formats_number: HashMap<String, ValueFormatNumber>,
    formats_percentage: HashMap<String, ValueFormatPercentage>,
    formats_currency: HashMap<String, ValueFormatCurrency>,
    formats_text: HashMap<String, ValueFormatText>,
    formats_datetime: HashMap<String, ValueFormatDateTime>,
    formats_timeduration: HashMap<String, ValueFormatTimeDuration>,

    /// Default-styles per Type.
    /// This is only used when writing the ods file.
    def_styles: HashMap<ValueType, String>,

    /// Page-layout data.
    pagestyles: HashMap<String, PageStyle>,
    masterpages: HashMap<String, MasterPage>,

    /// Validations.
    validations: HashMap<String, Validation>,

    /// Configuration data. Internal cache for all values.
    /// Mapped into WorkBookConfig, SheetConfig.
    config: Detach<Config>,
    /// User modifiable config.
    workbook_config: WorkBookConfig,
    /// Keeps all the namespaces.
    xmlns: HashMap<String, NamespaceMap>,

    /// All extra files contained in the zip manifest are copied here.
    manifest: HashMap<String, Manifest>,

    /// Metadata
    metadata: Metadata,

    /// other stuff ...
    extra: Vec<XmlTag>,
}

impl fmt::Debug for WorkBook {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
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
        for s in self.textstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.rubystyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.graphicstyles.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_boolean.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_number.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_percentage.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_currency.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_text.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_datetime.values() {
            writeln!(f, "{:?}", s)?;
        }
        for s in self.formats_timeduration.values() {
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
        for s in self.validations.values() {
            writeln!(f, "{:?}", s)?;
        }
        writeln!(f, "{:?}", &self.workbook_config)?;
        for v in self.manifest.values() {
            writeln!(f, "extras {:?}", v)?;
        }
        writeln!(f, "{:?}", &self.metadata)?;
        for xtr in &self.extra {
            writeln!(f, "extras {:?}", xtr)?;
        }
        Ok(())
    }
}

/// Autogenerate a stylename. Runs a counter with the prefix and
/// checks for existence.
fn auto_style_name<T>(
    autonum: &mut HashMap<String, u32>,
    prefix: &str,
    styles: &HashMap<String, T>,
) -> String {
    let mut cnt = if let Some(n) = autonum.get(prefix) {
        n + 1
    } else {
        0
    };

    let style_name = loop {
        let style_name = format!("{}{}", prefix, cnt);
        if !styles.contains_key(&style_name) {
            break style_name;
        }
        cnt += 1;
    };

    autonum.insert(prefix.to_string(), cnt);

    style_name
}

impl Default for WorkBook {
    fn default() -> Self {
        WorkBook::new(locale!("en"))
    }
}

impl WorkBook {
    /// Creates a new, completely empty workbook.
    ///
    /// WorkBook::locale_settings can be used to initialize default styles.
    pub fn new_empty() -> Self {
        WorkBook {
            sheets: Default::default(),
            version: "1.3".to_string(),
            fonts: Default::default(),
            autonum: Default::default(),
            scripts: Default::default(),
            event_listener: Default::default(),
            tablestyles: Default::default(),
            rowstyles: Default::default(),
            colstyles: Default::default(),
            cellstyles: Default::default(),
            paragraphstyles: Default::default(),
            textstyles: Default::default(),
            rubystyles: Default::default(),
            graphicstyles: Default::default(),
            formats_boolean: Default::default(),
            formats_number: Default::default(),
            formats_percentage: Default::default(),
            formats_currency: Default::default(),
            formats_text: Default::default(),
            formats_datetime: Default::default(),
            formats_timeduration: Default::default(),
            def_styles: Default::default(),
            pagestyles: Default::default(),
            masterpages: Default::default(),
            validations: Default::default(),
            config: default_settings(),
            workbook_config: Default::default(),
            extra: vec![],
            manifest: Default::default(),
            metadata: Default::default(),
            xmlns: Default::default(),
        }
    }

    /// Creates a new workbook, and initializes default styles according
    /// to the given locale.
    ///
    /// If the locale is not supported no ValueFormat's are set and all
    /// depends on the application opening the spreadsheet.
    ///
    /// The available locales can be activated via feature-flags.
    pub fn new(locale: Locale) -> Self {
        let mut wb = WorkBook::new_empty();
        wb.locale_settings(locale);
        wb
    }

    /// Creates a set of default formats and styles for every value-type.
    ///
    /// If the locale is not supported no ValueFormat's are set and all
    /// depends on the application opening the spreadsheet.
    ///
    /// The available locales can be activated via feature-flags.
    pub fn locale_settings(&mut self, locale: Locale) {
        if let Some(lf) = locale::localized_format(locale) {
            self.add_boolean_format(lf.boolean_format());
            self.add_number_format(lf.number_format());
            self.add_percentage_format(lf.percentage_format());
            self.add_currency_format(lf.currency_format());
            self.add_datetime_format(lf.date_format());
            self.add_datetime_format(lf.datetime_format());
            self.add_datetime_format(lf.time_of_day_format());
            self.add_timeduration_format(lf.time_interval_format());
        }

        self.add_cellstyle(CellStyle::new(
            DefaultStyle::bool().to_string(),
            &DefaultFormat::bool(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::number().to_string(),
            &DefaultFormat::number(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::percent().to_string(),
            &DefaultFormat::percent(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::currency().to_string(),
            &DefaultFormat::currency(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::date().to_string(),
            &DefaultFormat::date(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::datetime().to_string(),
            &DefaultFormat::datetime(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::time_of_day().to_string(),
            &DefaultFormat::time_of_day(),
        ));
        self.add_cellstyle(CellStyle::new(
            DefaultStyle::time_interval().to_string(),
            &DefaultFormat::time_interval(),
        ));

        self.add_def_style(ValueType::Boolean, &DefaultStyle::bool());
        self.add_def_style(ValueType::Number, &DefaultStyle::number());
        self.add_def_style(ValueType::Percentage, &DefaultStyle::percent());
        self.add_def_style(ValueType::Currency, &DefaultStyle::currency());
        self.add_def_style(ValueType::DateTime, &DefaultStyle::date());
        self.add_def_style(ValueType::TimeDuration, &DefaultStyle::time_interval());
    }

    /// ODS version. Defaults to 1.3.
    pub fn version(&self) -> &String {
        &self.version
    }

    /// ODS version. Defaults to 1.3.
    /// It's not advised to set another value.
    pub fn set_version(&mut self, version: String) {
        self.version = version;
    }

    /// Configuration flags.
    pub fn config(&self) -> &WorkBookConfig {
        &self.workbook_config
    }

    /// Configuration flags.
    pub fn config_mut(&mut self) -> &mut WorkBookConfig {
        &mut self.workbook_config
    }

    /// Number of sheets.
    pub fn num_sheets(&self) -> usize {
        self.sheets.len()
    }

    /// Finds the sheet index by the sheet-name.
    pub fn sheet_idx<S: AsRef<str>>(&self, name: S) -> Option<usize> {
        for (idx, sheet) in self.sheets.iter().enumerate() {
            if sheet.name == name.as_ref() {
                return Some(idx);
            }
        }
        None
    }

    /// Detaches a sheet.
    /// Useful if you have to make mutating calls to the workbook and
    /// the sheet intermixed.
    ///
    /// Warning
    ///
    /// The sheet has to be re-attached before saving the workbook.
    ///
    /// Panics
    ///
    /// Panics if the sheet has already been detached.
    /// Panics if n is out of bounds.
    pub fn detach_sheet(&mut self, n: usize) -> Detached<usize, Sheet> {
        self.sheets[n].detach(n)
    }

    /// Reattaches the sheet in the place it was before.
    ///
    /// Panics
    ///
    /// Panics if n is out of bounds.
    pub fn attach_sheet(&mut self, sheet: Detached<usize, Sheet>) {
        self.sheets[Detached::key(&sheet)].attach(sheet)
    }

    /// Returns a certain sheet.
    ///
    /// Panics
    ///
    /// Panics if n is out of bounds.
    pub fn sheet(&self, n: usize) -> &Sheet {
        self.sheets[n].as_ref()
    }

    /// Returns a certain sheet.
    ///
    /// Panics
    ///
    /// Panics if n does not exist.
    pub fn sheet_mut(&mut self, n: usize) -> &mut Sheet {
        self.sheets[n].as_mut()
    }

    /// Returns iterator over sheets.
    pub fn iter_sheets(&self) -> impl Iterator<Item = &Sheet> {
        self.sheets.iter().map(|sheet| &**sheet)
    }

    /// Inserts the sheet at the given position.
    pub fn insert_sheet(&mut self, i: usize, sheet: Sheet) {
        self.sheets.insert(i, sheet.into());
    }

    /// Appends a sheet.
    pub fn push_sheet(&mut self, sheet: Sheet) {
        self.sheets.push(sheet.into());
    }

    /// Removes a sheet from the table.
    ///
    /// Panics
    ///
    /// Panics if the sheet was detached.
    pub fn remove_sheet(&mut self, n: usize) -> Sheet {
        self.sheets.remove(n).take()
    }

    /// Scripts.
    pub fn add_script(&mut self, v: Script) {
        self.scripts.push(v);
    }

    /// Scripts.
    pub fn iter_scripts(&self) -> impl Iterator<Item = &Script> {
        self.scripts.iter()
    }

    /// Scripts
    pub fn scripts(&self) -> &Vec<Script> {
        &self.scripts
    }

    /// Scripts
    pub fn scripts_mut(&mut self) -> &mut Vec<Script> {
        &mut self.scripts
    }

    /// Event-Listener
    pub fn add_event_listener(&mut self, e: EventListener) {
        self.event_listener.insert(e.event_name.clone(), e);
    }

    /// Event-Listener
    pub fn remove_event_listener(&mut self, event_name: &str) -> Option<EventListener> {
        self.event_listener.remove(event_name)
    }

    /// Event-Listener
    pub fn iter_event_listeners(&self) -> impl Iterator<Item = &EventListener> {
        self.event_listener.values()
    }

    /// Event-Listener
    pub fn event_listener(&self, event_name: &str) -> Option<&EventListener> {
        self.event_listener.get(event_name)
    }

    /// Event-Listener
    pub fn event_listener_mut(&mut self, event_name: &str) -> Option<&mut EventListener> {
        self.event_listener.get_mut(event_name)
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

    /// Adds a font.
    pub fn add_font(&mut self, font: FontFaceDecl) {
        self.fonts.insert(font.name().to_string(), font);
    }

    /// Removes a font.
    pub fn remove_font(&mut self, name: &str) -> Option<FontFaceDecl> {
        self.fonts.remove(name)
    }

    /// Iterates the fonts.
    pub fn iter_fonts(&self) -> impl Iterator<Item = &FontFaceDecl> {
        self.fonts.values()
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
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_tablestyle(&mut self, mut style: TableStyle) -> TableStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(&mut self.autonum, "ta", &self.tablestyles));
        }
        let sref = style.style_ref();
        self.tablestyles.insert(style.name().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_tablestyle(&mut self, name: &str) -> Option<TableStyle> {
        self.tablestyles.remove(name)
    }

    /// Iterates the table-styles.
    pub fn iter_table_styles(&self) -> impl Iterator<Item = &TableStyle> {
        self.tablestyles.values()
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
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_rowstyle(&mut self, mut style: RowStyle) -> RowStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(&mut self.autonum, "ro", &self.rowstyles));
        }
        let sref = style.style_ref();
        self.rowstyles.insert(style.name().to_string(), style);
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

    /// Returns iterator over styles.
    pub fn iter_rowstyles(&self) -> impl Iterator<Item = &RowStyle> {
        self.rowstyles.values()
    }

    /// Adds a style.
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_colstyle(&mut self, mut style: ColStyle) -> ColStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(&mut self.autonum, "co", &self.colstyles));
        }
        let sref = style.style_ref();
        self.colstyles.insert(style.name().to_string(), style);
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

    /// Returns iterator over styles.
    pub fn iter_colstyles(&self) -> impl Iterator<Item = &ColStyle> {
        self.colstyles.values()
    }

    /// Adds a style.
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_cellstyle(&mut self, mut style: CellStyle) -> CellStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(&mut self.autonum, "ce", &self.cellstyles));
        }
        let sref = style.style_ref();
        self.cellstyles.insert(style.name().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_cellstyle(&mut self, name: &str) -> Option<CellStyle> {
        self.cellstyles.remove(name)
    }

    /// Returns iterator over styles.
    pub fn iter_cellstyles(&self) -> impl Iterator<Item = &CellStyle> {
        self.cellstyles.values()
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
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_paragraphstyle(&mut self, mut style: ParagraphStyle) -> ParagraphStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(
                &mut self.autonum,
                "para",
                &self.paragraphstyles,
            ));
        }
        let sref = style.style_ref();
        self.paragraphstyles.insert(style.name().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_paragraphstyle(&mut self, name: &str) -> Option<ParagraphStyle> {
        self.paragraphstyles.remove(name)
    }

    /// Returns iterator over styles.
    pub fn iter_paragraphstyles(&self) -> impl Iterator<Item = &ParagraphStyle> {
        self.paragraphstyles.values()
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
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_textstyle(&mut self, mut style: TextStyle) -> TextStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(&mut self.autonum, "txt", &self.textstyles));
        }
        let sref = style.style_ref();
        self.textstyles.insert(style.name().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_textstyle(&mut self, name: &str) -> Option<TextStyle> {
        self.textstyles.remove(name)
    }

    /// Returns iterator over styles.
    pub fn iter_textstyles(&self) -> impl Iterator<Item = &TextStyle> {
        self.textstyles.values()
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
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_rubystyle(&mut self, mut style: RubyStyle) -> RubyStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(&mut self.autonum, "ruby", &self.rubystyles));
        }
        let sref = style.style_ref();
        self.rubystyles.insert(style.name().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_rubystyle(&mut self, name: &str) -> Option<RubyStyle> {
        self.rubystyles.remove(name)
    }

    /// Returns iterator over styles.
    pub fn iter_rubystyles(&self) -> impl Iterator<Item = &RubyStyle> {
        self.rubystyles.values()
    }

    /// Returns the style.
    pub fn rubystyle(&self, name: &str) -> Option<&RubyStyle> {
        self.rubystyles.get(name)
    }

    /// Returns the mutable style.
    pub fn rubystyle_mut(&mut self, name: &str) -> Option<&mut RubyStyle> {
        self.rubystyles.get_mut(name)
    }

    /// Adds a style.
    /// Unnamed styles will be assigned an automatic name.
    pub fn add_graphicstyle(&mut self, mut style: GraphicStyle) -> GraphicStyleRef {
        if style.name().is_empty() {
            style.set_name(auto_style_name(
                &mut self.autonum,
                "gr",
                &self.graphicstyles,
            ));
        }
        let sref = style.style_ref();
        self.graphicstyles.insert(style.name().to_string(), style);
        sref
    }

    /// Removes a style.
    pub fn remove_graphicstyle(&mut self, name: &str) -> Option<GraphicStyle> {
        self.graphicstyles.remove(name)
    }

    /// Returns iterator over styles.
    pub fn iter_graphicstyles(&self) -> impl Iterator<Item = &GraphicStyle> {
        self.graphicstyles.values()
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
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_boolean_format(&mut self, mut vstyle: ValueFormatBoolean) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(&mut self.autonum, "val_boolean", &self.formats_boolean).as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_boolean
            .insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_boolean_format(&mut self, name: &str) -> Option<ValueFormatBoolean> {
        self.formats_boolean.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_boolean_formats(&self) -> impl Iterator<Item = &ValueFormatBoolean> {
        self.formats_boolean.values()
    }

    /// Returns the format.
    pub fn boolean_format(&self, name: &str) -> Option<&ValueFormatBoolean> {
        self.formats_boolean.get(name)
    }

    /// Returns the mutable format.
    pub fn boolean_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatBoolean> {
        self.formats_boolean.get_mut(name)
    }

    /// Adds a value format.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_number_format(&mut self, mut vstyle: ValueFormatNumber) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(&mut self.autonum, "val_number", &self.formats_number).as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_number
            .insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_number_format(&mut self, name: &str) -> Option<ValueFormatNumber> {
        self.formats_number.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_number_formats(&self) -> impl Iterator<Item = &ValueFormatNumber> {
        self.formats_number.values()
    }

    /// Returns the format.
    pub fn number_format(&self, name: &str) -> Option<&ValueFormatBoolean> {
        self.formats_boolean.get(name)
    }

    /// Returns the mutable format.
    pub fn number_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatBoolean> {
        self.formats_boolean.get_mut(name)
    }

    /// Adds a value format.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_percentage_format(&mut self, mut vstyle: ValueFormatPercentage) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(
                    &mut self.autonum,
                    "val_percentage",
                    &self.formats_percentage,
                )
                .as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_percentage
            .insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_percentage_format(&mut self, name: &str) -> Option<ValueFormatPercentage> {
        self.formats_percentage.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_percentage_formats(&self) -> impl Iterator<Item = &ValueFormatPercentage> {
        self.formats_percentage.values()
    }

    /// Returns the format.
    pub fn percentage_format(&self, name: &str) -> Option<&ValueFormatPercentage> {
        self.formats_percentage.get(name)
    }

    /// Returns the mutable format.
    pub fn percentage_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatPercentage> {
        self.formats_percentage.get_mut(name)
    }

    /// Adds a value format.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_currency_format(&mut self, mut vstyle: ValueFormatCurrency) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(&mut self.autonum, "val_currency", &self.formats_currency).as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_currency
            .insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_currency_format(&mut self, name: &str) -> Option<ValueFormatCurrency> {
        self.formats_currency.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_currency_formats(&self) -> impl Iterator<Item = &ValueFormatCurrency> {
        self.formats_currency.values()
    }

    /// Returns the format.
    pub fn currency_format(&self, name: &str) -> Option<&ValueFormatCurrency> {
        self.formats_currency.get(name)
    }

    /// Returns the mutable format.
    pub fn currency_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatCurrency> {
        self.formats_currency.get_mut(name)
    }

    /// Adds a value format.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_text_format(&mut self, mut vstyle: ValueFormatText) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(&mut self.autonum, "val_text", &self.formats_text).as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_text.insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_text_format(&mut self, name: &str) -> Option<ValueFormatText> {
        self.formats_text.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_text_formats(&self) -> impl Iterator<Item = &ValueFormatText> {
        self.formats_text.values()
    }

    /// Returns the format.
    pub fn text_format(&self, name: &str) -> Option<&ValueFormatText> {
        self.formats_text.get(name)
    }

    /// Returns the mutable format.
    pub fn text_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatText> {
        self.formats_text.get_mut(name)
    }

    /// Adds a value format.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_datetime_format(&mut self, mut vstyle: ValueFormatDateTime) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(&mut self.autonum, "val_datetime", &self.formats_datetime).as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_datetime
            .insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_datetime_format(&mut self, name: &str) -> Option<ValueFormatDateTime> {
        self.formats_datetime.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_datetime_formats(&self) -> impl Iterator<Item = &ValueFormatDateTime> {
        self.formats_datetime.values()
    }

    /// Returns the format.
    pub fn datetime_format(&self, name: &str) -> Option<&ValueFormatDateTime> {
        self.formats_datetime.get(name)
    }

    /// Returns the mutable format.
    pub fn datetime_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatDateTime> {
        self.formats_datetime.get_mut(name)
    }

    /// Adds a value format.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_timeduration_format(
        &mut self,
        mut vstyle: ValueFormatTimeDuration,
    ) -> ValueFormatRef {
        if vstyle.name().is_empty() {
            vstyle.set_name(
                auto_style_name(
                    &mut self.autonum,
                    "val_timeduration",
                    &self.formats_timeduration,
                )
                .as_str(),
            );
        }
        let sref = vstyle.format_ref();
        self.formats_timeduration
            .insert(vstyle.name().to_string(), vstyle);
        sref
    }

    /// Removes the format.
    pub fn remove_timeduration_format(&mut self, name: &str) -> Option<ValueFormatTimeDuration> {
        self.formats_timeduration.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_timeduration_formats(&self) -> impl Iterator<Item = &ValueFormatTimeDuration> {
        self.formats_timeduration.values()
    }

    /// Returns the format.
    pub fn timeduration_format(&self, name: &str) -> Option<&ValueFormatTimeDuration> {
        self.formats_timeduration.get(name)
    }

    /// Returns the mutable format.
    pub fn timeduration_format_mut(&mut self, name: &str) -> Option<&mut ValueFormatTimeDuration> {
        self.formats_timeduration.get_mut(name)
    }

    /// Adds a value PageStyle.
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_pagestyle(&mut self, mut pstyle: PageStyle) -> PageStyleRef {
        if pstyle.name().is_empty() {
            pstyle.set_name(auto_style_name(&mut self.autonum, "page", &self.pagestyles));
        }
        let sref = pstyle.style_ref();
        self.pagestyles.insert(pstyle.name().to_string(), pstyle);
        sref
    }

    /// Removes the PageStyle.
    pub fn remove_pagestyle(&mut self, name: &str) -> Option<PageStyle> {
        self.pagestyles.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_pagestyles(&self) -> impl Iterator<Item = &PageStyle> {
        self.pagestyles.values()
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
    /// Unnamed formats will be assigned an automatic name.
    pub fn add_masterpage(&mut self, mut mpage: MasterPage) -> MasterPageRef {
        if mpage.name().is_empty() {
            mpage.set_name(auto_style_name(&mut self.autonum, "mp", &self.masterpages));
        }
        let sref = mpage.masterpage_ref();
        self.masterpages.insert(mpage.name().to_string(), mpage);
        sref
    }

    /// Removes the MasterPage.
    pub fn remove_masterpage(&mut self, name: &str) -> Option<MasterPage> {
        self.masterpages.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_masterpages(&self) -> impl Iterator<Item = &MasterPage> {
        self.masterpages.values()
    }

    /// Returns the MasterPage.
    pub fn masterpage(&self, name: &str) -> Option<&MasterPage> {
        self.masterpages.get(name)
    }

    /// Returns the mutable MasterPage.
    pub fn masterpage_mut(&mut self, name: &str) -> Option<&mut MasterPage> {
        self.masterpages.get_mut(name)
    }

    /// Adds a Validation.
    /// Nameless validations will be assigned a name.
    pub fn add_validation(&mut self, mut valid: Validation) -> ValidationRef {
        if valid.name().is_empty() {
            valid.set_name(auto_style_name(&mut self.autonum, "val", &self.validations));
        }
        let vref = valid.validation_ref();
        self.validations.insert(valid.name().to_string(), valid);
        vref
    }

    /// Removes a Validation.
    pub fn remove_validation(&mut self, name: &str) -> Option<Validation> {
        self.validations.remove(name)
    }

    /// Returns iterator over formats.
    pub fn iter_validations(&self) -> impl Iterator<Item = &Validation> {
        self.validations.values()
    }

    /// Returns the Validation.
    pub fn validation(&self, name: &str) -> Option<&Validation> {
        self.validations.get(name)
    }

    /// Returns a mutable Validation.
    pub fn validation_mut(&mut self, name: &str) -> Option<&mut Validation> {
        self.validations.get_mut(name)
    }

    /// Adds a manifest entry, replaces an existing one with the same name.
    pub fn add_manifest(&mut self, manifest: Manifest) {
        self.manifest.insert(manifest.full_path.clone(), manifest);
    }

    /// Removes a manifest entry.
    pub fn remove_manifest(&mut self, path: &str) -> Option<Manifest> {
        self.manifest.remove(path)
    }

    /// Iterates the manifest.
    pub fn iter_manifest(&self) -> impl Iterator<Item = &Manifest> {
        self.manifest.values()
    }

    /// Returns the manifest entry for the path
    pub fn manifest(&self, path: &str) -> Option<&Manifest> {
        self.manifest.get(path)
    }

    /// Returns the manifest entry for the path
    pub fn manifest_mut(&mut self, path: &str) -> Option<&mut Manifest> {
        self.manifest.get_mut(path)
    }

    /// Gives access to meta-data.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Gives access to meta-data.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }
}

#[derive(Clone, Debug)]
pub(crate) struct NamespaceMap {
    map: HashMap<Cow<'static, str>, Cow<'static, str>>,
}

impl NamespaceMap {
    pub(crate) fn new() -> Self {
        Self {
            map: Default::default(),
        }
    }

    pub(crate) fn insert(&mut self, k: String, v: String) {
        self.map.insert(Cow::Owned(k), Cow::Owned(v));
    }

    pub(crate) fn insert_str(&mut self, k: &'static str, v: &'static str) {
        self.map.insert(Cow::Borrowed(k), Cow::Borrowed(v));
    }

    pub(crate) fn entries(&self) -> impl Iterator<Item = (&Cow<'static, str>, &Cow<'static, str>)> {
        self.map.iter()
    }
}

/// Subset of the Workbook wide configurations.
#[derive(Clone, Debug)]
pub struct WorkBookConfig {
    /// Which table is active when opening.    
    pub active_table: String,
    /// Show grid in general. Per sheet definition take priority.
    pub show_grid: bool,
    /// Show page-breaks.
    pub show_page_breaks: bool,
    /// Are the sheet-tabs shown or not.
    pub has_sheet_tabs: bool,
}

impl Default for WorkBookConfig {
    fn default() -> Self {
        Self {
            active_table: "".to_string(),
            show_grid: true,
            show_page_breaks: false,
            has_sheet_tabs: true,
        }
    }
}

/// Visibility of a column or row.
#[derive(Debug, Clone, Copy, Eq, PartialEq, Default)]
#[allow(missing_docs)]
pub enum Visibility {
    #[default]
    Visible,
    Collapsed,
    Filtered,
}

impl Display for Visibility {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), fmt::Error> {
        match self {
            Visibility::Visible => write!(f, "visible"),
            Visibility::Collapsed => write!(f, "collapse"),
            Visibility::Filtered => write!(f, "filter"),
        }
    }
}

/// Script.
#[derive(Debug, Default, Clone)]
pub struct Script {
    script_lang: String,
    script: Vec<XmlContent>,
}

impl Script {
    /// Script
    pub fn new() -> Self {
        Self {
            script_lang: "".to_string(),
            script: Default::default(),
        }
    }

    /// Script language
    pub fn script_lang(&self) -> &str {
        &self.script_lang
    }

    /// Script language
    pub fn set_script_lang(&mut self, script_lang: String) {
        self.script_lang = script_lang
    }

    /// Script
    pub fn script(&self) -> &Vec<XmlContent> {
        &self.script
    }

    /// Script
    pub fn set_script(&mut self, script: Vec<XmlContent>) {
        self.script = script
    }
}

/// Event-Listener.
#[derive(Debug, Clone)]
pub struct EventListener {
    event_name: String,
    script_lang: String,
    macro_name: String,
    actuate: XLinkActuate,
    href: String,
    link_type: XLinkType,
}

impl EventListener {
    /// EventListener
    pub fn new() -> Self {
        Self {
            event_name: Default::default(),
            script_lang: Default::default(),
            macro_name: Default::default(),
            actuate: XLinkActuate::OnLoad,
            href: Default::default(),
            link_type: Default::default(),
        }
    }

    /// Name
    pub fn event_name(&self) -> &str {
        &self.event_name
    }

    /// Name
    pub fn set_event_name(&mut self, name: String) {
        self.event_name = name;
    }

    /// Script language
    pub fn script_lang(&self) -> &str {
        &self.script_lang
    }

    /// Script language
    pub fn set_script_lang(&mut self, lang: String) {
        self.script_lang = lang
    }

    /// Macro name
    pub fn macro_name(&self) -> &str {
        &self.macro_name
    }

    /// Macro name
    pub fn set_macro_name(&mut self, name: String) {
        self.macro_name = name
    }

    /// Actuate
    pub fn actuate(&self) -> XLinkActuate {
        self.actuate
    }

    /// Actuate
    pub fn set_actuate(&mut self, actuate: XLinkActuate) {
        self.actuate = actuate;
    }

    /// HRef
    pub fn href(&self) -> &str {
        &self.href
    }

    /// HRef
    pub fn set_href(&mut self, href: String) {
        self.href = href;
    }

    /// Link type
    pub fn link_type(&self) -> XLinkType {
        self.link_type
    }

    /// Link type
    pub fn set_link_type(&mut self, link_type: XLinkType) {
        self.link_type = link_type
    }
}

impl Default for EventListener {
    fn default() -> Self {
        Self {
            event_name: Default::default(),
            script_lang: Default::default(),
            macro_name: Default::default(),
            actuate: XLinkActuate::OnRequest,
            href: Default::default(),
            link_type: Default::default(),
        }
    }
}

/// Row data
#[derive(Debug, Clone, Default)]
struct RowHeader {
    style: Option<String>,
    cellstyle: Option<String>,
    visible: Visibility,
    repeat: u32,
    height: Length,
}

impl RowHeader {
    pub(crate) fn new() -> Self {
        Self {
            style: None,
            cellstyle: None,
            visible: Default::default(),
            repeat: 1,
            height: Default::default(),
        }
    }

    pub(crate) fn set_style(&mut self, style: &RowStyleRef) {
        self.style = Some(style.to_string());
    }

    pub(crate) fn clear_style(&mut self) {
        self.style = None;
    }

    pub(crate) fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    pub(crate) fn set_cellstyle(&mut self, style: &CellStyleRef) {
        self.cellstyle = Some(style.to_string());
    }

    pub(crate) fn clear_cellstyle(&mut self) {
        self.cellstyle = None;
    }

    pub(crate) fn cellstyle(&self) -> Option<&String> {
        self.cellstyle.as_ref()
    }

    pub(crate) fn set_visible(&mut self, visible: Visibility) {
        self.visible = visible;
    }

    pub(crate) fn visible(&self) -> Visibility {
        self.visible
    }

    pub(crate) fn set_repeat(&mut self, repeat: u32) {
        assert!(repeat > 0);
        self.repeat = repeat;
    }

    pub(crate) fn repeat(&self) -> u32 {
        self.repeat
    }

    pub(crate) fn set_height(&mut self, height: Length) {
        self.height = height;
    }

    pub(crate) fn height(&self) -> Length {
        self.height
    }
}

/// Column data
#[derive(Debug, Clone, Default)]
struct ColHeader {
    style: Option<String>,
    cellstyle: Option<String>,
    visible: Visibility,
    width: Length,
}

impl ColHeader {
    pub(crate) fn new() -> Self {
        Self {
            style: None,
            cellstyle: None,
            visible: Default::default(),
            width: Default::default(),
        }
    }

    pub(crate) fn set_style(&mut self, style: &ColStyleRef) {
        self.style = Some(style.to_string());
    }

    pub(crate) fn clear_style(&mut self) {
        self.style = None;
    }

    pub(crate) fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    pub(crate) fn set_cellstyle(&mut self, style: &CellStyleRef) {
        self.cellstyle = Some(style.to_string());
    }

    pub(crate) fn clear_cellstyle(&mut self) {
        self.cellstyle = None;
    }

    pub(crate) fn cellstyle(&self) -> Option<&String> {
        self.cellstyle.as_ref()
    }

    pub(crate) fn set_visible(&mut self, visible: Visibility) {
        self.visible = visible;
    }

    pub(crate) fn visible(&self) -> Visibility {
        self.visible
    }

    pub(crate) fn set_width(&mut self, width: Length) {
        self.width = width;
    }

    pub(crate) fn width(&self) -> Length {
        self.width
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

    data: BTreeMap<(u32, u32), CellData>,

    col_header: BTreeMap<u32, ColHeader>,
    row_header: BTreeMap<u32, RowHeader>,

    display: bool,
    print: bool,

    print_ranges: Option<Vec<CellRange>>,

    group_rows: Vec<Grouped>,
    group_cols: Vec<Grouped>,

    sheet_config: SheetConfig,

    extra: Vec<XmlTag>,
}

impl<'a> IntoIterator for &'a Sheet {
    type Item = ((u32, u32), CellContentRef<'a>);
    type IntoIter = CellIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        CellIter {
            it_data: self.data.iter(),
            k_data: None,
            v_data: None,
        }
    }
}

/// Iterator over cells.
#[derive(Clone, Debug)]
pub struct CellIter<'a> {
    it_data: std::collections::btree_map::Iter<'a, (u32, u32), CellData>,
    k_data: Option<&'a (u32, u32)>,
    v_data: Option<&'a CellData>,
}

impl CellIter<'_> {
    /// Returns the (row,col) of the next cell.
    pub fn peek_cell(&mut self) -> Option<(u32, u32)> {
        self.k_data.copied()
    }

    fn load_next_data(&mut self) {
        if let Some((k, v)) = self.it_data.next() {
            self.k_data = Some(k);
            self.v_data = Some(v);
        } else {
            self.k_data = None;
            self.v_data = None;
        }
    }
}

impl FusedIterator for CellIter<'_> {}

impl<'a> Iterator for CellIter<'a> {
    type Item = ((u32, u32), CellContentRef<'a>);

    fn next(&mut self) -> Option<Self::Item> {
        if self.k_data.is_none() {
            self.load_next_data();
        }

        if let Some(k_data) = self.k_data {
            if let Some(v_data) = self.v_data {
                let r = Some((*k_data, v_data.cell_content_ref()));
                self.load_next_data();
                r
            } else {
                None
            }
        } else {
            None
        }
    }
}

/// Range iterator.
#[derive(Clone, Debug)]
pub struct Range<'a> {
    range: std::collections::btree_map::Range<'a, (u32, u32), CellData>,
}

impl FusedIterator for Range<'_> {}

impl<'a> Iterator for Range<'a> {
    type Item = ((u32, u32), CellContentRef<'a>);

    fn next(&mut self) -> Option<Self::Item> {
        if let Some((k, v)) = self.range.next() {
            Some((*k, v.cell_content_ref()))
        } else {
            None
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.range.size_hint()
    }
}

impl DoubleEndedIterator for Range<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        if let Some((k, v)) = self.range.next_back() {
            Some((*k, v.cell_content_ref()))
        } else {
            None
        }
    }
}

impl ExactSizeIterator for Range<'_> {}

impl fmt::Debug for Sheet {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
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
        for v in &self.group_cols {
            writeln!(f, "group cols {:?}", v)?;
        }
        for v in &self.group_rows {
            writeln!(f, "group rows {:?}", v)?;
        }
        for xtr in &self.extra {
            writeln!(f, "extras {:?}", xtr)?;
        }
        Ok(())
    }
}

impl Sheet {
    /// Create an empty sheet.
    #[deprecated]
    pub fn new_with_name<S: Into<String>>(name: S) -> Self {
        Self::new(name)
    }

    /// Create an empty sheet.
    ///
    /// The name is shown as the tab-title, but also as a reference for
    /// this sheet in formulas and sheet-metadata. Giving an empty string
    /// here is allowed and a name will be generated, when the document is
    /// opened. But any metadata will not be applied.
    ///
    /// Renaming the sheet works for metadata, but formulas will not be fixed.  
    ///
    pub fn new<S: Into<String>>(name: S) -> Self {
        Sheet {
            name: name.into(),
            data: BTreeMap::new(),
            col_header: Default::default(),
            style: None,
            print_ranges: None,
            group_rows: Default::default(),
            group_cols: Default::default(),
            sheet_config: Default::default(),
            extra: vec![],
            row_header: Default::default(),
            display: true,
            print: true,
        }
    }

    /// Copy all the attributes but not the actual data.
    pub fn clone_no_data(&self) -> Self {
        Self {
            name: self.name.clone(),
            style: self.style.clone(),
            data: Default::default(),
            col_header: self.col_header.clone(),
            row_header: self.row_header.clone(),
            display: self.display,
            print: self.print,
            print_ranges: self.print_ranges.clone(),
            group_rows: self.group_rows.clone(),
            group_cols: self.group_cols.clone(),
            sheet_config: Default::default(),
            extra: self.extra.clone(),
        }
    }

    /// Iterate all cells.
    pub fn iter(&self) -> CellIter<'_> {
        self.into_iter()
    }

    /// Count all cells with any data.
    pub fn cell_count(&self) -> usize {
        self.data.len()
    }

    /// Iterate a range of cells.
    pub fn range<R>(&self, range: R) -> Range<'_>
    where
        R: RangeBounds<(u32, u32)>,
    {
        Range {
            range: self.data.range(range),
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

    /// Configuration for the sheet.
    pub fn config(&self) -> &SheetConfig {
        &self.sheet_config
    }

    /// Configuration for the sheet.
    pub fn config_mut(&mut self) -> &mut SheetConfig {
        &mut self.sheet_config
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
    pub fn set_colstyle(&mut self, col: u32, style: &ColStyleRef) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_style(style);
    }

    /// Remove the style.
    pub fn clear_colstyle(&mut self, col: u32) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .clear_style();
    }

    /// Returns the column style.
    pub fn colstyle(&self, col: u32) -> Option<&String> {
        if let Some(col_header) = self.col_header.get(&col) {
            col_header.style()
        } else {
            None
        }
    }

    /// Default cell style for this column.
    pub fn set_col_cellstyle(&mut self, col: u32, style: &CellStyleRef) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_cellstyle(style);
    }

    /// Remove the style.
    pub fn clear_col_cellstyle(&mut self, col: u32) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .clear_cellstyle();
    }

    /// Returns the default cell style for this column.
    pub fn col_cellstyle(&self, col: u32) -> Option<&String> {
        if let Some(col_header) = self.col_header.get(&col) {
            col_header.cellstyle()
        } else {
            None
        }
    }

    /// Visibility of the column
    pub fn set_col_visible(&mut self, col: u32, visible: Visibility) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_visible(visible);
    }

    /// Returns the default cell style for this column.
    pub fn col_visible(&self, col: u32) -> Visibility {
        if let Some(col_header) = self.col_header.get(&col) {
            col_header.visible()
        } else {
            Default::default()
        }
    }

    /// Sets the column width for this column.
    pub fn set_col_width(&mut self, col: u32, width: Length) {
        self.col_header
            .entry(col)
            .or_insert_with(ColHeader::new)
            .set_width(width);
    }

    /// Returns the column-width.
    pub fn col_width(&self, col: u32) -> Length {
        if let Some(ch) = self.col_header.get(&col) {
            ch.width()
        } else {
            Length::Default
        }
    }

    /// Row style.
    pub fn set_rowstyle(&mut self, row: u32, style: &RowStyleRef) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .set_style(style);
    }

    /// Remove the style.
    pub fn clear_rowstyle(&mut self, row: u32) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .clear_style();
    }

    /// Returns the row style.
    pub fn rowstyle(&self, row: u32) -> Option<&String> {
        if let Some(row_header) = self.row_header.get(&row) {
            row_header.style()
        } else {
            None
        }
    }

    /// Default cell style for this row.
    pub fn set_row_cellstyle(&mut self, row: u32, style: &CellStyleRef) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .set_cellstyle(style);
    }

    /// Remove the style.
    pub fn clear_row_cellstyle(&mut self, row: u32) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .clear_cellstyle();
    }

    /// Returns the default cell style for this row.
    pub fn row_cellstyle(&self, row: u32) -> Option<&String> {
        if let Some(row_header) = self.row_header.get(&row) {
            row_header.cellstyle()
        } else {
            None
        }
    }

    /// Visibility of the row
    pub fn set_row_visible(&mut self, row: u32, visible: Visibility) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .set_visible(visible);
    }

    /// Returns the default cell style for this row.
    pub fn row_visible(&self, row: u32) -> Visibility {
        if let Some(row_header) = self.row_header.get(&row) {
            row_header.visible()
        } else {
            Default::default()
        }
    }

    /// Sets the repeat count for this row. Usually this is the last row
    /// with data in a sheet. Setting the repeat count will not change
    /// the row number of following rows. But they will be changed after
    /// writing to an ODS file and reading it again.
    ///
    /// Panics
    ///
    /// Panics if the repeat is 0.
    pub fn set_row_repeat(&mut self, row: u32, repeat: u32) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .set_repeat(repeat)
    }

    /// Returns the repeat count for this row.
    pub fn row_repeat(&self, row: u32) -> u32 {
        if let Some(row_header) = self.row_header.get(&row) {
            row_header.repeat()
        } else {
            Default::default()
        }
    }

    /// Sets the row-height.
    pub fn set_row_height(&mut self, row: u32, height: Length) {
        self.row_header
            .entry(row)
            .or_insert_with(RowHeader::new)
            .set_height(height);
    }

    /// Returns the row-height
    pub fn row_height(&self, row: u32) -> Length {
        if let Some(rh) = self.row_header.get(&row) {
            rh.height()
        } else {
            Length::Default
        }
    }

    /// Returns the maximum used column +1 in the column header
    pub fn used_cols(&self) -> u32 {
        *self.col_header.keys().max().unwrap_or(&0) + 1
    }

    /// Returns the maximum used row +1 in the row header
    pub fn used_rows(&self) -> u32 {
        *self.row_header.keys().max().unwrap_or(&0) + 1
    }

    /// Returns a tuple of (max(row)+1, max(col)+1)
    pub fn used_grid_size(&self) -> (u32, u32) {
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
    pub fn is_empty(&self, row: u32, col: u32) -> bool {
        self.data.get(&(row, col)).is_none()
    }

    /// Basic range operator.
    // pub fn cell_range<R>(&self, range: R)
    // where
    //     R: RangeBounds<(ucell, ucell)>,
    // {
    //     let r = self.data.range(range);
    // }

    /// Returns a copy of the cell content.
    pub fn cell(&self, row: u32, col: u32) -> Option<CellContent> {
        self.data
            .get(&(row, col))
            .map(CellData::cloned_cell_content)
    }

    /// Returns a copy of the cell content.
    pub fn cell_ref(&self, row: u32, col: u32) -> Option<CellContentRef<'_>> {
        self.data.get(&(row, col)).map(CellData::cell_content_ref)
    }

    /// Consumes the CellContent and sets the values.
    pub fn add_cell(&mut self, row: u32, col: u32, cell: CellContent) {
        self.add_cell_data(row, col, cell.into_celldata());
    }

    /// Removes the cell and returns the values as CellContent.
    pub fn remove_cell(&mut self, row: u32, col: u32) -> Option<CellContent> {
        self.data
            .remove(&(row, col))
            .map(CellData::into_cell_content)
    }

    /// Add a new cell. Main use is for reading the spreadsheet.
    pub(crate) fn add_cell_data(&mut self, row: u32, col: u32, cell: CellData) {
        self.data.insert((row, col), cell);
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_styled_value<V: Into<Value>>(
        &mut self,
        row: u32,
        col: u32,
        value: V,
        style: &CellStyleRef,
    ) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.value = value.into();
        cell.style = Some(style.to_string());
    }

    /// Sets a value for the specified cell. Creates a new cell if necessary.
    pub fn set_value<V: Into<Value>>(&mut self, row: u32, col: u32, value: V) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.value = value.into();
    }

    /// Returns a value
    pub fn value(&self, row: u32, col: u32) -> &Value {
        if let Some(cell) = self.data.get(&(row, col)) {
            &cell.value
        } else {
            &Value::Empty
        }
    }

    /// Sets a formula for the specified cell. Creates a new cell if necessary.
    pub fn set_formula<V: Into<String>>(&mut self, row: u32, col: u32, formula: V) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.formula = Some(formula.into());
    }

    /// Removes the formula.
    pub fn clear_formula(&mut self, row: u32, col: u32) {
        if let Some(cell) = self.data.get_mut(&(row, col)) {
            cell.formula = None;
        }
    }

    /// Returns a value
    pub fn formula(&self, row: u32, col: u32) -> Option<&String> {
        if let Some(c) = self.data.get(&(row, col)) {
            c.formula.as_ref()
        } else {
            None
        }
    }

    /// Sets a repeat counter for the cell.
    pub fn set_col_repeat(&mut self, row: u32, col: u32, repeat: u32) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.repeat = repeat;
    }

    /// Returns the repeat counter for the cell.
    pub fn col_repeat(&self, row: u32, col: u32) -> u32 {
        if let Some(c) = self.data.get(&(row, col)) {
            c.repeat
        } else {
            1
        }
    }

    /// Sets the cell-style for the specified cell. Creates a new cell if necessary.
    pub fn set_cellstyle(&mut self, row: u32, col: u32, style: &CellStyleRef) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.style = Some(style.to_string());
    }

    /// Removes the cell-style.
    pub fn clear_cellstyle(&mut self, row: u32, col: u32) {
        if let Some(cell) = self.data.get_mut(&(row, col)) {
            cell.style = None;
        }
    }

    /// Returns a value
    pub fn cellstyle(&self, row: u32, col: u32) -> Option<&String> {
        if let Some(c) = self.data.get(&(row, col)) {
            c.style.as_ref()
        } else {
            None
        }
    }

    /// Sets a content-validation for this cell.
    pub fn set_validation(&mut self, row: u32, col: u32, validation: &ValidationRef) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().validation_name = Some(validation.to_string());
    }

    /// Removes the cell-style.
    pub fn clear_validation(&mut self, row: u32, col: u32) {
        if let Some(cell) = self.data.get_mut(&(row, col)) {
            cell.extra_mut().validation_name = None;
        }
    }

    /// Returns a content-validation name for this cell.
    pub fn validation(&self, row: u32, col: u32) -> Option<&String> {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            c.validation_name.as_ref()
        } else {
            None
        }
    }

    /// Sets the rowspan of the cell. Must be greater than 0.
    pub fn set_row_span(&mut self, row: u32, col: u32, span: u32) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().span.row_span = span;
    }

    /// Rowspan of the cell.
    pub fn row_span(&self, row: u32, col: u32) -> u32 {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            c.span.row_span
        } else {
            1
        }
    }

    /// Sets the colspan of the cell. Must be greater than 0.
    pub fn set_col_span(&mut self, row: u32, col: u32, span: u32) {
        assert!(span > 0);
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().span.col_span = span;
    }

    /// Colspan of the cell.
    pub fn col_span(&self, row: u32, col: u32) -> u32 {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            c.span.col_span
        } else {
            1
        }
    }

    /// Sets the rowspan of the cell. Must be greater than 0.
    pub fn set_matrix_row_span(&mut self, row: u32, col: u32, span: u32) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().matrix_span.row_span = span;
    }

    /// Rowspan of the cell.
    pub fn matrix_row_span(&self, row: u32, col: u32) -> u32 {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            c.matrix_span.row_span
        } else {
            1
        }
    }

    /// Sets the colspan of the cell. Must be greater than 0.
    pub fn set_matrix_col_span(&mut self, row: u32, col: u32, span: u32) {
        assert!(span > 0);
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().matrix_span.col_span = span;
    }

    /// Colspan of the cell.
    pub fn matrix_col_span(&self, row: u32, col: u32) -> u32 {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            c.matrix_span.col_span
        } else {
            1
        }
    }

    /// Sets a annotation for this cell.
    pub fn set_annotation(&mut self, row: u32, col: u32, annotation: Annotation) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().annotation = Some(annotation);
    }

    /// Removes the annotation.
    pub fn clear_annotation(&mut self, row: u32, col: u32) {
        if let Some(cell) = self.data.get_mut(&(row, col)) {
            cell.extra_mut().annotation = None;
        }
    }

    /// Returns a content-validation name for this cell.
    pub fn annotation(&self, row: u32, col: u32) -> Option<&Annotation> {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            c.annotation.as_ref()
        } else {
            None
        }
    }

    /// Returns a content-validation name for this cell.
    pub fn annotation_mut(&mut self, row: u32, col: u32) -> Option<&mut Annotation> {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get_mut(&(row, col)) {
            c.annotation.as_mut()
        } else {
            None
        }
    }

    /// Add a drawframe to a specific cell.
    pub fn add_draw_frame(&mut self, row: u32, col: u32, draw_frame: DrawFrame) {
        let cell = self
            .data
            .entry((row, col))
            .or_insert_with(CellData::default);
        cell.extra_mut().draw_frames.push(draw_frame);
    }

    /// Removes all drawframes.
    pub fn clear_draw_frames(&mut self, row: u32, col: u32) {
        if let Some(cell) = self.data.get_mut(&(row, col)) {
            cell.extra_mut().draw_frames = Vec::new();
        }
    }

    /// Returns the draw-frames.
    pub fn draw_frames(&self, row: u32, col: u32) -> Option<&Vec<DrawFrame>> {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get(&(row, col)) {
            Some(c.draw_frames.as_ref())
        } else {
            None
        }
    }

    /// Returns a content-validation name for this cell.
    pub fn draw_frames_mut(&mut self, row: u32, col: u32) -> Option<&mut Vec<DrawFrame>> {
        if let Some(CellData { extra: Some(c), .. }) = self.data.get_mut(&(row, col)) {
            Some(c.draw_frames.as_mut())
        } else {
            None
        }
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

    /// Split horizontally on a cell boundary. The splitting is fixed in
    /// position.
    pub fn split_col_header(&mut self, col: u32) {
        self.config_mut().hor_split_mode = SplitMode::Heading;
        self.config_mut().hor_split_pos = col + 1;
        self.config_mut().position_right = col + 1;
        self.config_mut().cursor_x = col + 1;
    }

    /// Split vertically on a cell boundary. The splitting is fixed in
    /// position.
    pub fn split_row_header(&mut self, row: u32) {
        self.config_mut().vert_split_mode = SplitMode::Heading;
        self.config_mut().vert_split_pos = row + 1;
        self.config_mut().position_bottom = row + 1;
        self.config_mut().cursor_y = row + 1;
    }

    /// Split horizontally with a pixel width. The split can be moved around.
    /// For more control look at SheetConfig.
    pub fn split_horizontal(&mut self, col: u32) {
        self.config_mut().hor_split_mode = SplitMode::Split;
        self.config_mut().hor_split_pos = col;
    }

    /// Split vertically with a pixel width. The split can be moved around.
    /// For more control look at SheetConfig.
    pub fn split_vertical(&mut self, col: u32) {
        self.config_mut().vert_split_mode = SplitMode::Split;
        self.config_mut().vert_split_pos = col;
    }

    /// Add a column group.
    ///
    /// Panic
    ///
    /// Column groups can be contained within another, but they can't overlap.
    /// From must be less than or equal to.
    pub fn add_col_group(&mut self, from: u32, to: u32) {
        assert!(from <= to);
        let grp = Grouped::new(from, to, true);
        for v in &self.group_cols {
            assert!(grp.contains(v) || v.contains(&grp) || grp.disjunct(v));
        }
        self.group_cols.push(grp);
    }

    /// Remove a column group.
    pub fn remove_col_group(&mut self, from: u32, to: u32) {
        if let Some(idx) = self
            .group_cols
            .iter()
            .position(|v| v.from == from && v.to == to)
        {
            self.group_cols.remove(idx);
        }
    }

    /// Change the expansion/collapse of a col group.
    ///
    /// Does nothing if no such group exists.
    pub fn set_col_group_displayed(&mut self, from: u32, to: u32, display: bool) {
        if let Some(idx) = self
            .group_cols
            .iter()
            .position(|v| v.from == from && v.to == to)
        {
            self.group_cols[idx].display = display;

            for r in from..=to {
                self.set_col_visible(
                    r,
                    if display {
                        Visibility::Visible
                    } else {
                        Visibility::Collapsed
                    },
                );
            }
        }
    }

    /// Count of column groups.
    pub fn col_group_count(&self) -> usize {
        self.group_cols.len()
    }

    /// Returns the nth column group.
    pub fn col_group(&self, idx: usize) -> Option<&Grouped> {
        return self.group_cols.get(idx);
    }

    /// Returns the nth column group.
    pub fn col_group_mut(&mut self, idx: usize) -> Option<&mut Grouped> {
        return self.group_cols.get_mut(idx);
    }

    /// Iterate the column groups.
    pub fn col_group_iter(&self) -> impl Iterator<Item = &Grouped> {
        self.group_cols.iter()
    }

    /// Add a row group.
    ///
    /// Panic
    ///
    /// Row groups can be contained within another, but they can't overlap.
    /// From must be less than or equal to.
    pub fn add_row_group(&mut self, from: u32, to: u32) {
        assert!(from <= to);
        let grp = Grouped::new(from, to, true);
        for v in &self.group_rows {
            assert!(grp.contains(v) || v.contains(&grp) || grp.disjunct(v));
        }
        self.group_rows.push(grp);
    }

    /// Remove a row group.
    pub fn remove_row_group(&mut self, from: u32, to: u32) {
        if let Some(idx) = self
            .group_rows
            .iter()
            .position(|v| v.from == from && v.to == to)
        {
            self.group_rows.remove(idx);
        }
    }

    /// Change the expansion/collapse of a row group.
    ///
    /// Does nothing if no such group exists.
    pub fn set_row_group_displayed(&mut self, from: u32, to: u32, display: bool) {
        if let Some(idx) = self
            .group_rows
            .iter()
            .position(|v| v.from == from && v.to == to)
        {
            self.group_rows[idx].display = display;

            for r in from..=to {
                self.set_row_visible(
                    r,
                    if display {
                        Visibility::Visible
                    } else {
                        Visibility::Collapsed
                    },
                );
            }
        }
    }

    /// Count of row groups.
    pub fn row_group_count(&self) -> usize {
        self.group_rows.len()
    }

    /// Returns the nth row group.
    pub fn row_group(&self, idx: usize) -> Option<&Grouped> {
        return self.group_rows.get(idx);
    }

    /// Iterate row groups.
    pub fn row_group_iter(&self) -> impl Iterator<Item = &Grouped> {
        self.group_rows.iter()
    }
}

/// Describes a row/column group.
#[derive(Debug, PartialEq, Clone)]
pub struct Grouped {
    /// Inclusive from row/col.
    pub from: u32,
    /// Inclusive to row/col.
    pub to: u32,
    /// Visible/Collapsed state.
    pub display: bool,
}

impl Grouped {
    /// New group.
    pub fn new(from: u32, to: u32, display: bool) -> Self {
        Self { from, to, display }
    }

    /// Inclusive start.
    pub fn from(&self) -> u32 {
        self.from
    }

    /// Inclusive start.
    pub fn set_from(&mut self, from: u32) {
        self.from = from;
    }

    /// Inclusive end.
    pub fn to(&self) -> u32 {
        self.to
    }

    /// Inclusive end.
    pub fn set_to(&mut self, to: u32) {
        self.to = to
    }

    /// Group is displayed?
    pub fn display(&self) -> bool {
        self.display
    }

    /// Change the display state for the group.
    ///
    /// Note: Changing this does not change the visibility of the rows/columns.
    /// Use Sheet::set_display_col_group() and Sheet::set_display_row_group() to make
    /// all necessary changes.
    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }

    /// Contains the other group.
    pub fn contains(&self, other: &Grouped) -> bool {
        self.from <= other.from && self.to >= other.to
    }

    /// The two groups are disjunct.
    pub fn disjunct(&self, other: &Grouped) -> bool {
        self.from < other.from && self.to < other.from || self.from > other.to && self.to > other.to
    }
}

/// There are two ways a sheet can be split. There are fixed column/row header
/// like splits, and there is a moveable split.
///
#[derive(Clone, Copy, Debug)]
#[allow(missing_docs)]
pub enum SplitMode {
    None = 0,
    Split = 1,
    Heading = 2,
}

impl TryFrom<i16> for SplitMode {
    type Error = OdsError;

    fn try_from(n: i16) -> Result<Self, Self::Error> {
        match n {
            0 => Ok(SplitMode::None),
            1 => Ok(SplitMode::Split),
            2 => Ok(SplitMode::Heading),
            _ => Err(OdsError::Ods(format!("Invalid split mode {}", n))),
        }
    }
}

/// Per sheet configurations.
#[derive(Clone, Debug)]
pub struct SheetConfig {
    /// Active column.
    pub cursor_x: u32,
    /// Active row.
    pub cursor_y: u32,
    /// Splitting the table.
    pub hor_split_mode: SplitMode,
    /// Splitting the table.
    pub vert_split_mode: SplitMode,
    /// Position of the split.
    pub hor_split_pos: u32,
    /// Position of the split.
    pub vert_split_pos: u32,
    /// SplitMode is Pixel
    /// - 0-4 indicates the quadrant where the focus is.
    /// SplitMode is Cell
    /// - No real function.
    pub active_split_range: i16,
    /// SplitMode is Pixel
    /// - First visible column in the left quadrant.
    /// SplitMode is Cell
    /// - The first visible column in the left quadrant.
    ///   AND every column left of this one is simply invisible.
    pub position_left: u32,
    /// SplitMode is Pixel
    /// - First visible column in the right quadrant.
    /// SplitMode is Cell
    /// - The first visible column in the right quadrant.
    pub position_right: u32,
    /// SplitMode is Pixel
    /// - First visible row in the top quadrant.
    /// SplitMode is Cell
    /// - The first visible row in the top quadrant.
    ///   AND every row up from this one is simply invisible.
    pub position_top: u32,
    /// SplitMode is Pixel
    /// - The first visible row in teh right quadrant.
    /// SplitMode is Cell
    /// - The first visible row in the bottom quadrant.
    pub position_bottom: u32,
    /// If 0 then zoom_value denotes a percentage.
    /// If 2 then zoom_value is 50%???
    pub zoom_type: i16,
    /// Value of zoom.
    pub zoom_value: i32,
    /// Value of pageview zoom.
    pub page_view_zoom_value: i32,
    /// Grid is showing.
    pub show_grid: bool,
}

impl Default for SheetConfig {
    fn default() -> Self {
        Self {
            cursor_x: 0,
            cursor_y: 0,
            hor_split_mode: SplitMode::None,
            vert_split_mode: SplitMode::None,
            hor_split_pos: 0,
            vert_split_pos: 0,
            active_split_range: 2,
            position_left: 0,
            position_right: 0,
            position_top: 0,
            position_bottom: 0,
            zoom_type: 0,
            zoom_value: 100,
            page_view_zoom_value: 60,
            show_grid: true,
        }
    }
}

/// A cell can span multiple rows/columns.
#[derive(Debug, Clone, Copy)]
pub struct CellSpan {
    row_span: u32,
    col_span: u32,
}

impl Default for CellSpan {
    fn default() -> Self {
        Self {
            row_span: 1,
            col_span: 1,
        }
    }
}

impl From<CellSpan> for (u32, u32) {
    fn from(span: CellSpan) -> Self {
        (span.row_span, span.col_span)
    }
}

impl From<&CellSpan> for (u32, u32) {
    fn from(span: &CellSpan) -> Self {
        (span.row_span, span.col_span)
    }
}

impl CellSpan {
    /// Default span 1,1
    pub fn new() -> Self {
        Self::default()
    }

    /// Is this empty? Defined as row_span==1 and col_span==1.
    pub fn is_empty(&self) -> bool {
        self.row_span == 1 && self.col_span == 1
    }

    /// Sets the row span of this cell.
    /// Cells below with values will be lost when writing.
    pub fn set_row_span(&mut self, rows: u32) {
        assert!(rows > 0);
        self.row_span = rows;
    }

    /// Returns the row span.
    pub fn row_span(&self) -> u32 {
        self.row_span
    }

    /// Sets the column span of this cell.
    /// Cells to the right with values will be lost when writing.
    pub fn set_col_span(&mut self, cols: u32) {
        assert!(cols > 0);
        self.col_span = cols;
    }

    /// Returns the col span.
    pub fn col_span(&self) -> u32 {
        self.col_span
    }
}

/// One Cell of the spreadsheet.
#[derive(Debug, Clone)]
struct CellData {
    value: Value,
    // Unparsed formula string.
    formula: Option<String>,
    // Cell style name.
    style: Option<String>,
    // Cell repeated.
    repeat: u32,
    // Scarcely used extra data.
    extra: Option<Box<CellDataExt>>,
}

/// Extra cell data.
#[derive(Debug, Clone, Default)]
struct CellDataExt {
    // Content validation name.
    validation_name: Option<String>,
    // Row/Column span.
    span: CellSpan,
    // Matrix span.
    matrix_span: CellSpan,
    // Annotation
    annotation: Option<Annotation>,
    // Draw
    draw_frames: Vec<DrawFrame>,
}

impl Default for CellData {
    fn default() -> Self {
        Self {
            value: Default::default(),
            formula: None,
            style: None,
            repeat: 1,
            extra: None,
        }
    }
}

impl CellData {
    pub(crate) fn extra_mut(&mut self) -> &mut CellDataExt {
        if self.extra.is_none() {
            self.extra = Some(Box::default());
        }
        self.extra.as_mut().expect("celldataext")
    }

    pub(crate) fn cloned_cell_content(&self) -> CellContent {
        let (validation_name, span, matrix_span, annotation, draw_frames) =
            if let Some(extra) = &self.extra {
                (
                    extra.validation_name.clone(),
                    extra.span,
                    extra.matrix_span,
                    extra.annotation.clone(),
                    extra.draw_frames.clone(),
                )
            } else {
                (
                    None,
                    Default::default(),
                    Default::default(),
                    None,
                    Vec::new(),
                )
            };

        CellContent {
            value: self.value.clone(),
            style: self.style.clone(),
            formula: self.formula.clone(),
            repeat: self.repeat,
            validation_name,
            span,
            matrix_span,
            annotation,
            draw_frames,
        }
    }

    pub(crate) fn into_cell_content(self) -> CellContent {
        let (validation_name, span, matrix_span, annotation, draw_frames) =
            if let Some(extra) = self.extra {
                (
                    extra.validation_name,
                    extra.span,
                    extra.matrix_span,
                    extra.annotation,
                    extra.draw_frames,
                )
            } else {
                (
                    None,
                    Default::default(),
                    Default::default(),
                    None,
                    Vec::new(),
                )
            };

        CellContent {
            value: self.value,
            style: self.style,
            formula: self.formula,
            repeat: self.repeat,
            validation_name,
            span,
            matrix_span,
            annotation,
            draw_frames,
        }
    }

    pub(crate) fn cell_content_ref(&self) -> CellContentRef<'_> {
        let (validation_name, span, matrix_span, annotation, draw_frames) =
            if let Some(extra) = &self.extra {
                (
                    extra.validation_name.as_ref(),
                    Some(&extra.span),
                    Some(&extra.matrix_span),
                    extra.annotation.as_ref(),
                    Some(&extra.draw_frames),
                )
            } else {
                (None, None, None, None, None)
            };

        CellContentRef {
            value: &self.value,
            style: self.style.as_ref(),
            formula: self.formula.as_ref(),
            repeat: &self.repeat,
            validation_name,
            span,
            matrix_span,
            annotation,
            draw_frames,
        }
    }
}

/// Holds references to the combined content of a cell.
/// A temporary to hold the data when iterating over a sheet.
#[derive(Debug, Clone, Copy)]
pub struct CellContentRef<'a> {
    /// Reference to the cell value.
    pub value: &'a Value,
    /// Reference to the stylename.
    pub style: Option<&'a String>,
    /// Reference to the cell formula.
    pub formula: Option<&'a String>,
    /// Reference to the repeat count.
    pub repeat: &'a u32,
    /// Reference to a cell validation.
    pub validation_name: Option<&'a String>,
    /// Reference to the cellspan.
    pub span: Option<&'a CellSpan>,
    /// Reference to a matrix cellspan.
    pub matrix_span: Option<&'a CellSpan>,
    /// Reference to an annotation.
    pub annotation: Option<&'a Annotation>,
    /// Reference to draw-frames.
    pub draw_frames: Option<&'a Vec<DrawFrame>>,
}

impl<'a> CellContentRef<'a> {
    /// Returns the value.
    pub fn value(&self) -> &'a Value {
        self.value
    }

    /// Returns the formula.
    pub fn formula(&self) -> Option<&'a String> {
        self.formula
    }

    /// Returns the cell style.
    pub fn style(&self) -> Option<&'a String> {
        self.style
    }

    /// Returns the repeat count.
    pub fn repeat(&self) -> &'a u32 {
        self.repeat
    }

    /// Returns the validation name.
    pub fn validation(&self) -> Option<&'a String> {
        self.validation_name
    }

    /// Returns the row span.
    pub fn row_span(&self) -> u32 {
        if let Some(span) = self.span {
            span.row_span
        } else {
            1
        }
    }

    /// Returns the col span.
    pub fn col_span(&self) -> u32 {
        if let Some(span) = self.span {
            span.col_span
        } else {
            1
        }
    }

    /// Returns the row span for a matrix.
    pub fn matrix_row_span(&self) -> u32 {
        if let Some(matrix_span) = self.matrix_span {
            matrix_span.row_span
        } else {
            1
        }
    }

    /// Returns the col span for a matrix.
    pub fn matrix_col_span(&self) -> u32 {
        if let Some(matrix_span) = self.matrix_span {
            matrix_span.col_span
        } else {
            1
        }
    }

    /// Returns the validation name.
    pub fn annotation(&self) -> Option<&'a Annotation> {
        self.annotation
    }

    /// Returns draw frames.
    pub fn draw_frames(&self) -> Option<&'a Vec<DrawFrame>> {
        self.draw_frames
    }
}

/// A copy of the relevant data for a spreadsheet cell.
#[derive(Debug, Clone, Default)]
pub struct CellContent {
    /// Cell value.
    pub value: Value,
    /// Cell stylename.
    pub style: Option<String>,
    /// Cell formula.
    pub formula: Option<String>,
    /// Cell repeat count.
    pub repeat: u32,
    /// Reference to a validation rule.
    pub validation_name: Option<String>,
    /// Cellspan.
    pub span: CellSpan,
    /// Matrix span.
    pub matrix_span: CellSpan,
    /// Annotation
    pub annotation: Option<Annotation>,
    /// DrawFrames
    pub draw_frames: Vec<DrawFrame>,
}

impl CellContent {
    /// Empty.
    pub fn new() -> Self {
        Default::default()
    }

    ///
    pub(crate) fn into_celldata(mut self) -> CellData {
        let extra = self.into_celldata_ext();
        CellData {
            value: self.value,
            formula: self.formula,
            style: self.style,
            repeat: self.repeat,
            extra,
        }
    }

    /// Move stuff into a CellDataExt.
    #[allow(clippy::wrong_self_convention)]
    pub(crate) fn into_celldata_ext(&mut self) -> Option<Box<CellDataExt>> {
        if self.validation_name.is_some()
            || !self.span.is_empty()
            || !self.matrix_span.is_empty()
            || self.annotation.is_some()
            || !self.draw_frames.is_empty()
        {
            Some(Box::new(CellDataExt {
                validation_name: self.validation_name.take(),
                span: self.span,
                matrix_span: self.matrix_span,
                annotation: self.annotation.take(),
                draw_frames: std::mem::take(&mut self.draw_frames),
            }))
        } else {
            None
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

    /// Resets the formula.
    pub fn clear_formula(&mut self) {
        self.formula = None;
    }

    /// Returns the cell style.
    pub fn style(&self) -> Option<&String> {
        self.style.as_ref()
    }

    /// Sets the cell style.
    pub fn set_style(&mut self, style: &CellStyleRef) {
        self.style = Some(style.to_string());
    }

    /// Removes the style.
    pub fn clear_style(&mut self) {
        self.style = None;
    }

    /// Sets the repeat count for the cell.
    /// Value must be > 0.
    pub fn set_repeat(&mut self, repeat: u32) {
        assert!(repeat > 0);
        self.repeat = repeat;
    }

    /// Returns the repeat count for the cell.
    pub fn get_repeat(&mut self) -> u32 {
        self.repeat
    }

    /// Returns the validation name.
    pub fn validation(&self) -> Option<&String> {
        self.validation_name.as_ref()
    }

    /// Sets the validation name.
    pub fn set_validation(&mut self, validation: &ValidationRef) {
        self.validation_name = Some(validation.to_string());
    }

    /// No validation.
    pub fn clear_validation(&mut self) {
        self.validation_name = None;
    }

    /// Sets the row span of this cell.
    /// Cells below with values will be lost when writing.
    pub fn set_row_span(&mut self, rows: u32) {
        assert!(rows > 0);
        self.span.row_span = rows;
    }

    /// Returns the row span.
    pub fn row_span(&self) -> u32 {
        self.span.row_span
    }

    /// Sets the column span of this cell.
    /// Cells to the right with values will be lost when writing.
    pub fn set_col_span(&mut self, cols: u32) {
        assert!(cols > 0);
        self.span.col_span = cols;
    }

    /// Returns the col span.
    pub fn col_span(&self) -> u32 {
        self.span.col_span
    }

    /// Sets the row span of this cell.
    /// Cells below with values will be lost when writing.
    pub fn set_matrix_row_span(&mut self, rows: u32) {
        assert!(rows > 0);
        self.matrix_span.row_span = rows;
    }

    /// Returns the row span.
    pub fn matrix_row_span(&self) -> u32 {
        self.matrix_span.row_span
    }

    /// Sets the column span of this cell.
    /// Cells to the right with values will be lost when writing.
    pub fn set_matrix_col_span(&mut self, cols: u32) {
        assert!(cols > 0);
        self.matrix_span.col_span = cols;
    }

    /// Returns the col span.
    pub fn matrix_col_span(&self) -> u32 {
        self.matrix_span.col_span
    }

    /// Annotation
    pub fn set_annotation(&mut self, annotation: Annotation) {
        self.annotation = Some(annotation);
    }

    /// Annotation
    pub fn clear_annotation(&mut self) {
        self.annotation = None;
    }

    /// Returns the Annotation
    pub fn annotation(&self) -> Option<&Annotation> {
        self.annotation.as_ref()
    }

    /// Draw Frames
    pub fn set_draw_frames(&mut self, draw_frames: Vec<DrawFrame>) {
        self.draw_frames = draw_frames;
    }

    /// Draw Frames
    pub fn draw_frames(&self) -> &Vec<DrawFrame> {
        &self.draw_frames
    }
}

/// Datatypes for the values. Only the discriminants of the Value enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[allow(missing_docs)]
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

/// Content-Values
#[derive(Debug, Clone, PartialEq, Default)]
#[allow(missing_docs)]
pub enum Value {
    #[default]
    Empty,
    Boolean(bool),
    Number(f64),
    Percentage(f64),
    Currency(f64, String),
    Text(String),
    TextXml(Vec<TextTag>),
    DateTime(NaiveDateTime),
    TimeDuration(Duration),
}

impl Value {
    /// Return the plan ValueType for this value.
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

    /// Return the content as i64 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_i64_or(&self, d: i64) -> i64 {
        match self {
            Value::Number(n) => *n as i64,
            Value::Percentage(p) => *p as i64,
            Value::Currency(v, _) => *v as i64,
            _ => d,
        }
    }

    /// Return the content as i64 if the value is a number, percentage or
    /// currency.
    pub fn as_i64_opt(&self) -> Option<i64> {
        match self {
            Value::Number(n) => Some(*n as i64),
            Value::Percentage(p) => Some(*p as i64),
            Value::Currency(v, _) => Some(*v as i64),
            _ => None,
        }
    }

    /// Return the content as u64 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_u64_or(&self, d: u64) -> u64 {
        match self {
            Value::Number(n) => *n as u64,
            Value::Percentage(p) => *p as u64,
            Value::Currency(v, _) => *v as u64,
            _ => d,
        }
    }

    /// Return the content as u64 if the value is a number, percentage or
    /// currency.
    pub fn as_u64_opt(&self) -> Option<u64> {
        match self {
            Value::Number(n) => Some(*n as u64),
            Value::Percentage(p) => Some(*p as u64),
            Value::Currency(v, _) => Some(*v as u64),
            _ => None,
        }
    }

    /// Return the content as i32 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_i32_or(&self, d: i32) -> i32 {
        match self {
            Value::Number(n) => *n as i32,
            Value::Percentage(p) => *p as i32,
            Value::Currency(v, _) => *v as i32,
            _ => d,
        }
    }

    /// Return the content as i32 if the value is a number, percentage or
    /// currency.
    pub fn as_i32_opt(&self) -> Option<i32> {
        match self {
            Value::Number(n) => Some(*n as i32),
            Value::Percentage(p) => Some(*p as i32),
            Value::Currency(v, _) => Some(*v as i32),
            _ => None,
        }
    }

    /// Return the content as u32 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_u32_or(&self, d: u32) -> u32 {
        match self {
            Value::Number(n) => *n as u32,
            Value::Percentage(p) => *p as u32,
            Value::Currency(v, _) => *v as u32,
            _ => d,
        }
    }

    /// Return the content as u32 if the value is a number, percentage or
    /// currency.
    pub fn as_u32_opt(&self) -> Option<u32> {
        match self {
            Value::Number(n) => Some(*n as u32),
            Value::Percentage(p) => Some(*p as u32),
            Value::Currency(v, _) => Some(*v as u32),
            _ => None,
        }
    }

    /// Return the content as i16 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_i16_or(&self, d: i16) -> i16 {
        match self {
            Value::Number(n) => *n as i16,
            Value::Percentage(p) => *p as i16,
            Value::Currency(v, _) => *v as i16,
            _ => d,
        }
    }

    /// Return the content as i16 if the value is a number, percentage or
    /// currency.
    pub fn as_i16_opt(&self) -> Option<i16> {
        match self {
            Value::Number(n) => Some(*n as i16),
            Value::Percentage(p) => Some(*p as i16),
            Value::Currency(v, _) => Some(*v as i16),
            _ => None,
        }
    }

    /// Return the content as u16 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_u16_or(&self, d: u16) -> u16 {
        match self {
            Value::Number(n) => *n as u16,
            Value::Percentage(p) => *p as u16,
            Value::Currency(v, _) => *v as u16,
            _ => d,
        }
    }

    /// Return the content as u16 if the value is a number, percentage or
    /// currency.
    pub fn as_u16_opt(&self) -> Option<u16> {
        match self {
            Value::Number(n) => Some(*n as u16),
            Value::Percentage(p) => Some(*p as u16),
            Value::Currency(v, _) => Some(*v as u16),
            _ => None,
        }
    }

    /// Return the content as i8 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_i8_or(&self, d: i8) -> i8 {
        match self {
            Value::Number(n) => *n as i8,
            Value::Percentage(p) => *p as i8,
            Value::Currency(v, _) => *v as i8,
            _ => d,
        }
    }

    /// Return the content as i8 if the value is a number, percentage or
    /// currency.
    pub fn as_i8_opt(&self) -> Option<i8> {
        match self {
            Value::Number(n) => Some(*n as i8),
            Value::Percentage(p) => Some(*p as i8),
            Value::Currency(v, _) => Some(*v as i8),
            _ => None,
        }
    }

    /// Return the content as u8 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_u8_or(&self, d: u8) -> u8 {
        match self {
            Value::Number(n) => *n as u8,
            Value::Percentage(p) => *p as u8,
            Value::Currency(v, _) => *v as u8,
            _ => d,
        }
    }

    /// Return the content as u8 if the value is a number, percentage or
    /// currency.
    pub fn as_u8_opt(&self) -> Option<u8> {
        match self {
            Value::Number(n) => Some(*n as u8),
            Value::Percentage(p) => Some(*p as u8),
            Value::Currency(v, _) => Some(*v as u8),
            _ => None,
        }
    }

    /// Return the content as decimal if the value is a number, percentage or
    /// currency. Default otherwise.
    #[cfg(feature = "use_decimal")]
    pub fn as_decimal_or(&self, d: Decimal) -> Decimal {
        match self {
            Value::Number(n) => Decimal::from_f64(*n).unwrap_or(d),
            Value::Currency(v, _) => Decimal::from_f64(*v).unwrap_or(d),
            Value::Percentage(p) => Decimal::from_f64(*p).unwrap_or(d),
            _ => d,
        }
    }

    /// Return the content as decimal if the value is a number, percentage or
    /// currency. Default otherwise.
    #[cfg(feature = "use_decimal")]
    pub fn as_decimal_opt(&self) -> Option<Decimal> {
        match self {
            Value::Number(n) => Decimal::from_f64(*n),
            Value::Currency(v, _) => Decimal::from_f64(*v),
            Value::Percentage(p) => Decimal::from_f64(*p),
            _ => None,
        }
    }

    /// Return the content as f64 if the value is a number, percentage or
    /// currency. Default otherwise.
    pub fn as_f64_or(&self, d: f64) -> f64 {
        match self {
            Value::Number(n) => *n,
            Value::Currency(v, _) => *v,
            Value::Percentage(p) => *p,
            _ => d,
        }
    }

    /// Return the content as f64 if the value is a number, percentage or
    /// currency.
    pub fn as_f64_opt(&self) -> Option<f64> {
        match self {
            Value::Number(n) => Some(*n),
            Value::Currency(v, _) => Some(*v),
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

    /// Return the content as str if the value is text or markup text.
    /// When the cell contains markup all the markup is removed, but
    /// line-breaks are kept as \n.
    pub fn as_cow_str_or<'a>(&'a self, d: &'a str) -> Cow<'a, str> {
        match self {
            Value::Text(s) => Cow::from(s),
            Value::TextXml(v) => {
                let mut buf = String::new();
                for t in v {
                    if !buf.is_empty() {
                        buf.push('\n');
                    }
                    t.extract_text(&mut buf);
                }
                Cow::from(buf)
            }
            _ => Cow::from(d),
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

    /// Return the content as NaiveDate if the value is a DateTime.
    /// Default otherwise.
    pub fn as_date_or(&self, d: NaiveDate) -> NaiveDate {
        match self {
            Value::DateTime(dt) => dt.date(),
            _ => d,
        }
    }

    /// Return the content as an optional NaiveDateTime if the value is
    /// a DateTime.
    pub fn as_date_opt(&self) -> Option<NaiveDate> {
        match self {
            Value::DateTime(dt) => Some(dt.date()),
            _ => None,
        }
    }

    /// Returns the currency code or "" if the value is not a currency.
    pub fn currency(&self) -> &str {
        match self {
            Value::Currency(_, c) => c,
            _ => "",
        }
    }

    /// Create a currency value.
    #[allow(clippy::needless_range_loop)]
    pub fn new_currency<S: AsRef<str>>(cur: S, value: f64) -> Self {
        Value::Currency(value, cur.as_ref().to_string())
    }

    /// Create a percentage value.
    pub fn new_percentage(value: f64) -> Self {
        Value::Percentage(value)
    }
}

/// currency value
#[macro_export]
macro_rules! currency {
    ($c:expr, $v:expr) => {
        Value::new_currency($c, $v as f64)
    };
}

/// currency value
#[macro_export]
macro_rules! percent {
    ($v:expr) => {
        Value::new_percentage($v)
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
        Value::TextXml(vec![t])
    }
}

impl From<Vec<TextTag>> for Value {
    fn from(t: Vec<TextTag>) -> Self {
        Value::TextXml(t)
    }
}

impl From<Option<&str>> for Value {
    fn from(s: Option<&str>) -> Self {
        if let Some(s) = s {
            Value::Text(s.to_string())
        } else {
            Value::Empty
        }
    }
}

impl From<Option<&String>> for Value {
    fn from(s: Option<&String>) -> Self {
        if let Some(s) = s {
            Value::Text(s.to_string())
        } else {
            Value::Empty
        }
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
        Value::Number(f.to_f64().expect("decimal->f64 should not fail"))
    }
}

#[cfg(feature = "use_decimal")]
impl From<Option<Decimal>> for Value {
    fn from(f: Option<Decimal>) -> Self {
        if let Some(f) = f {
            Value::Number(f.to_f64().expect("decimal->f64 should not fail"))
        } else {
            Value::Empty
        }
    }
}

macro_rules! from_number {
    ($l:ty) => {
        impl From<$l> for Value {
            #![allow(trivial_numeric_casts)]
            fn from(f: $l) -> Self {
                Value::Number(f as f64)
            }
        }

        impl From<&$l> for Value {
            #![allow(trivial_numeric_casts)]
            fn from(f: &$l) -> Self {
                Value::Number(*f as f64)
            }
        }

        impl From<Option<$l>> for Value {
            #![allow(trivial_numeric_casts)]
            fn from(f: Option<$l>) -> Self {
                if let Some(f) = f {
                    Value::Number(f as f64)
                } else {
                    Value::Empty
                }
            }
        }

        impl From<Option<&$l>> for Value {
            #![allow(trivial_numeric_casts)]
            fn from(f: Option<&$l>) -> Self {
                if let Some(f) = f {
                    Value::Number(*f as f64)
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
        Value::DateTime(dt.and_hms_opt(0, 0, 0).unwrap())
    }
}

impl From<Option<NaiveDate>> for Value {
    fn from(dt: Option<NaiveDate>) -> Self {
        if let Some(dt) = dt {
            Value::DateTime(dt.and_hms_opt(0, 0, 0).expect("valid time"))
        } else {
            Value::Empty
        }
    }
}

impl From<NaiveTime> for Value {
    fn from(ti: NaiveTime) -> Self {
        Value::DateTime(NaiveDateTime::new(
            NaiveDate::from_ymd_opt(1900, 1, 1).expect("valid date"),
            ti,
        ))
    }
}

impl From<Option<NaiveTime>> for Value {
    fn from(dt: Option<NaiveTime>) -> Self {
        if let Some(ti) = dt {
            Value::DateTime(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(1900, 1, 1).expect("valid date"),
                ti,
            ))
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
