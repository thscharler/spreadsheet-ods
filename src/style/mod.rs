//! Styles define a large number of attributes. These are grouped together
//! as table, row, column, cell, paragraph and text attributes.
//!
//! ```
//! use spreadsheet_ods::{Style, CellRef, WorkBook};
//! use spreadsheet_ods::style::{StyleOrigin, StyleUse, AttrText, StyleMap, TableCellStyle};
//! use color::Rgb;
//!
//! let mut wb = WorkBook::new();
//!
//! let mut st = TableCellStyle::new("ce12", "num2");
//! st.set_color(Rgb::new(192, 128, 0));
//! st.set_font_bold();
//! wb.add_cell_style(st);
//!
//! let mut st = TableCellStyle::new("ce11", "num2");
//! st.set_color(Rgb::new(0, 192, 128));
//! st.set_font_bold();
//! wb.add_cell_style(st);
//!
//! let mut st = TableCellStyle::new("ce13", "num4");
//! st.push_stylemap(StyleMap::new("cell-content()=\"BB\"", "ce12", CellRef::remote("sheet0", 4, 3)));
//! st.push_stylemap(StyleMap::new("cell-content()=\"CC\"", "ce11", CellRef::remote("sheet0", 4, 3)));
//! wb.add_cell_style(st);
//! ```
//! Styles can be defined in content.xml or as global styles in styles.xml. This
//! is reflected as the StyleOrigin. The StyleUse differentiates between automatic
//! and user visible, named styles. And third StyleFor defines for which part of
//! the document the style can be used.
//!
//! Cell styles usually reference a value format for text formatting purposes.
//!
//! Styles can also link to a parent style and to a pagelayout.
//!

mod attr;
#[macro_use]
mod attr_macro;
mod cell_style;
mod column_style;
mod fontface;
mod graphic_style;
mod pagelayout;
mod paragraph_style;
mod row_style;
mod stylemap;
mod table_style;
mod tabstop;
mod text_style;
mod units;

pub use crate::attrmap::*;
pub use attr::*;
pub use cell_style::*;
pub use column_style::*;
pub use fontface::*;
pub use graphic_style::*;
pub use pagelayout::*;
pub use paragraph_style::*;
pub use row_style::*;
pub use stylemap::*;
pub use table_style::*;
pub use tabstop::*;
pub use text_style::*;
pub use units::*;

use crate::sealed::Sealed;
use color::Rgb;
use string_cache::DefaultAtom;

/// Origin of a style. Content.xml or Styles.xml.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleOrigin {
    Content,
    Styles,
}

impl Default for StyleOrigin {
    fn default() -> Self {
        StyleOrigin::Content
    }
}

/// Placement of a style. office:styles or office:automatic-styles
/// Defines the usage pattern for the style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleUse {
    Default,
    Named,
    Automatic,
}

impl Default for StyleUse {
    fn default() -> Self {
        StyleUse::Automatic
    }
}

/// Applicability of this style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleFor {
    Table,
    TableRow,
    TableColumn,
    TableCell,
    Graphic,
    Paragraph,
    Text,
    None,
}

impl Default for StyleFor {
    fn default() -> Self {
        StyleFor::None
    }
}

#[derive(Debug, Clone, Default)]
pub struct Style {
    /// Style name.
    name: String,
    /// Nice String.
    display_name: Option<String>,
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// Applicability of this style.
    family: StyleFor,
    /// Styles can cascade.
    parent: Option<String>,
    /// References the actual formatting instructions in the value-styles.
    value_format: Option<String>,
    /// References a page format. Only valid for table styles.
    master_page_name: Option<String>,
    /// Table styling
    table_attr: TableAttr,
    /// Column styling
    table_col_attr: TableColAttr,
    /// Row styling
    table_row_attr: TableRowAttr,
    /// Cell styles
    table_cell_attr: TableCellAttr,
    /// Text paragraph styles
    paragraph_attr: ParagraphAttr,
    /// Text styles
    text_attr: TextAttr,
    /// Graphic styles
    graphic_attr: GraphicAttr,
    /// Style maps
    stylemaps: Option<Vec<StyleMap>>,
}

impl Style {
    /// New, empty.
    pub fn new() -> Self {
        Style {
            name: String::from(""),
            display_name: None,
            origin: Default::default(),
            styleuse: Default::default(),
            family: Default::default(),
            parent: None,
            value_format: None,
            master_page_name: None,
            table_attr: Default::default(),
            table_col_attr: Default::default(),
            table_row_attr: Default::default(),
            table_cell_attr: Default::default(),
            paragraph_attr: Default::default(),
            text_attr: Default::default(),
            graphic_attr: Default::default(),
            stylemaps: None,
        }
    }

    /// Creates a new cell style.
    /// value_style references a ValueFormat.
    pub fn new_cell_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::new_with_name(StyleFor::TableCell, name, value_style)
    }

    /// Creates a new column style.
    /// value_style references a ValueFormat.
    pub fn new_col_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::new_with_name(StyleFor::TableColumn, name, value_style)
    }

    /// Creates a new row style.
    /// value_style references a ValueFormat.
    pub fn new_row_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::new_with_name(StyleFor::TableRow, name, value_style)
    }

    /// Creates a new table style.
    /// value_style references a ValueFormat.
    pub fn new_table_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::new_with_name(StyleFor::Table, name, value_style)
    }

    /// New, with name.
    /// value_style references a ValueFormat.
    pub fn new_with_name<S: Into<String>, T: Into<String>>(
        family: StyleFor,
        name: S,
        value_style: T,
    ) -> Self {
        Style {
            name: name.into(),
            display_name: None,
            origin: Default::default(),
            styleuse: Default::default(),
            family,
            parent: Some(String::from("Default")),
            value_format: Some(value_style.into()),
            master_page_name: None,
            table_attr: Default::default(),
            table_col_attr: Default::default(),
            table_row_attr: Default::default(),
            table_cell_attr: Default::default(),
            paragraph_attr: Default::default(),
            text_attr: Default::default(),
            graphic_attr: Default::default(),
            stylemaps: None,
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

    /// Sets the display name.
    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.display_name = Some(name.into());
    }

    /// Returns the display name.
    pub fn display_name(&self) -> Option<&String> {
        self.display_name.as_ref()
    }

    /// Sets the origin.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Returns the origin.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Style usage.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Returns the usage.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Sets the style-family.
    pub fn set_family(&mut self, family: StyleFor) {
        self.family = family;
    }

    /// Returns the style-family.
    pub fn family(&self) -> StyleFor {
        self.family
    }

    /// Sets the parent style.
    pub fn set_parent<S: Into<String>>(&mut self, parent: S) {
        self.parent = Some(parent.into());
    }

    /// Returns the parent style.
    pub fn parent(&self) -> Option<&String> {
        self.parent.as_ref()
    }

    /// Sets the value format.
    pub fn set_value_format<S: Into<String>>(&mut self, value_format: S) {
        self.value_format = Some(value_format.into());
    }

    /// Returns the value format.
    pub fn value_format(&self) -> Option<&String> {
        self.value_format.as_ref()
    }

    /// Sets the value format.
    pub fn set_master_page_name<S: Into<String>>(&mut self, value_format: S) {
        self.master_page_name = Some(value_format.into());
    }

    /// Returns the value format.
    pub fn master_page_name(&self) -> Option<&String> {
        self.master_page_name.as_ref()
    }

    /// Table style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::Table.
    pub fn table(&self) -> &TableAttr {
        assert_eq!(
            self.family,
            StyleFor::Table,
            "Can only be used for Table-Style."
        );
        &self.table_attr
    }

    /// Table style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::Table.    
    pub fn table_mut(&mut self) -> &mut TableAttr {
        assert_eq!(
            self.family,
            StyleFor::Table,
            "Can only be used for Table-Style."
        );
        &mut self.table_attr
    }

    /// Table column style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableColumn.
    pub fn col(&self) -> &TableColAttr {
        assert_eq!(
            self.family,
            StyleFor::TableColumn,
            "Can only be used for Column-Style."
        );
        &self.table_col_attr
    }

    /// Table column style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableColumn.
    pub fn col_mut(&mut self) -> &mut TableColAttr {
        assert_eq!(
            self.family,
            StyleFor::TableColumn,
            "Can only be used for Column-Style."
        );
        &mut self.table_col_attr
    }

    /// Table-row style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableRow.
    pub fn row(&self) -> &TableRowAttr {
        assert_eq!(
            self.family,
            StyleFor::TableRow,
            "Can only be used for Row-Style."
        );
        &self.table_row_attr
    }

    /// Table-row style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableRow.
    pub fn row_mut(&mut self) -> &mut TableRowAttr {
        assert_eq!(
            self.family,
            StyleFor::TableRow,
            "Can only be used for Row-Style."
        );
        &mut self.table_row_attr
    }

    /// Table-cell style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableCell.    
    pub fn cell(&self) -> &TableCellAttr {
        assert_eq!(
            self.family,
            StyleFor::TableCell,
            "Can only be used for Cell-Style."
        );
        &self.table_cell_attr
    }

    /// Table-cell style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableCell.    
    pub fn cell_mut(&mut self) -> &mut TableCellAttr {
        assert_eq!(
            self.family,
            StyleFor::TableCell,
            "Can only be used for Cell-Style."
        );
        &mut self.table_cell_attr
    }

    /// Paragraph style attributes.
    pub fn paragraph(&self) -> &ParagraphAttr {
        &self.paragraph_attr
    }

    /// Paragraph style attributes.
    pub fn paragraph_mut(&mut self) -> &mut ParagraphAttr {
        &mut self.paragraph_attr
    }

    /// Graphic style attributes.
    pub fn graphic(&self) -> &GraphicAttr {
        &self.graphic_attr
    }

    /// Graphic style attributes.
    pub fn graphic_mut(&mut self) -> &mut GraphicAttr {
        &mut self.graphic_attr
    }

    /// Text style attributes.
    pub fn text(&self) -> &TextAttr {
        &self.text_attr
    }

    /// Text style attributes.
    pub fn text_mut(&mut self) -> &mut TextAttr {
        &mut self.text_attr
    }

    /// Adds a stylemap.
    pub fn push_stylemap(&mut self, stylemap: StyleMap) {
        self.stylemaps.get_or_insert_with(Vec::new).push(stylemap);
    }

    /// Returns the stylemaps
    pub fn stylemaps(&self) -> Option<&Vec<StyleMap>> {
        self.stylemaps.as_ref()
    }

    /// Returns the mutable stylemap.
    pub fn stylemaps_mut(&mut self) -> &mut Vec<StyleMap> {
        self.stylemaps.get_or_insert_with(Vec::new)
    }
}

/// Style for the whole table.
#[derive(Clone, Debug, Default)]
pub struct TableAttr {
    attr: AttrMapType,
}

impl Sealed for TableAttr {}

impl AttrMap for TableAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a TableAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBackgroundColor for TableAttr {}

impl AttrFoMargin for TableAttr {}

impl AttrFoBreak for TableAttr {}

impl AttrFoKeepWithNext for TableAttr {}

impl AttrStyleShadow for TableAttr {}

impl AttrStyleWritingMode for TableAttr {}

/// Styles for table rows.
#[derive(Clone, Debug, Default)]
pub struct TableRowAttr {
    attr: AttrMapType,
}

impl Sealed for TableRowAttr {}

impl AttrMap for TableRowAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a TableRowAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBackgroundColor for TableRowAttr {}

impl AttrFoBreak for TableRowAttr {}

impl AttrFoKeepTogether for TableRowAttr {}

impl AttrTableRow for TableRowAttr {}

/// Styles for table columns.
#[derive(Clone, Debug, Default)]
pub struct TableColAttr {
    attr: AttrMapType,
}

impl Sealed for TableColAttr {}

impl AttrMap for TableColAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a TableColAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBreak for TableColAttr {}

impl AttrTableCol for TableColAttr {}

/// Styles for table cells.
#[derive(Clone, Debug, Default)]
pub struct TableCellAttr {
    attr: AttrMapType,
}

impl Sealed for TableCellAttr {}

impl AttrMap for TableCellAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a TableCellAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBackgroundColor for TableCellAttr {}

impl AttrFoBorder for TableCellAttr {}

impl AttrFoPadding for TableCellAttr {}

impl AttrStyleShadow for TableCellAttr {}

impl AttrStyleWritingMode for TableCellAttr {}

impl AttrTableCell for TableCellAttr {}

impl<'a> IntoIterator for &'a TabStop {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

/// Paragraph style.
#[derive(Clone, Debug, Default)]
pub struct ParagraphAttr {
    attr: AttrMapType,
    tabstops: Option<Vec<TabStop>>,
    // todo: drop-cap
    // todo: background-image
}

impl ParagraphAttr {
    pub fn add_tabstop(&mut self, ts: TabStop) {
        let tabstops = self.tabstops.get_or_insert_with(Vec::new);
        tabstops.push(ts);
    }

    pub fn tabstops(&self) -> Option<&Vec<TabStop>> {
        self.tabstops.as_ref()
    }
}

impl Sealed for ParagraphAttr {}

impl AttrMap for ParagraphAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a ParagraphAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBackgroundColor for ParagraphAttr {}

impl AttrFoBorder for ParagraphAttr {}

impl AttrFoBreak for ParagraphAttr {}

impl AttrFoKeepTogether for ParagraphAttr {}

impl AttrFoKeepWithNext for ParagraphAttr {}

impl AttrFoMargin for ParagraphAttr {}

impl AttrFoPadding for ParagraphAttr {}

impl AttrStyleShadow for ParagraphAttr {}

impl AttrStyleWritingMode for ParagraphAttr {}

impl AttrParagraph for ParagraphAttr {}

/// Text styles.
#[derive(Clone, Debug, Default)]
pub struct TextAttr {
    attr: AttrMapType,
}

impl Sealed for TextAttr {}

impl AttrMap for TextAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a TextAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBackgroundColor for TextAttr {}

impl AttrText for TextAttr {}

#[derive(Clone, Debug, Default)]
pub struct GraphicAttr {
    attr: AttrMapType,
}

impl Sealed for GraphicAttr {}

impl AttrMap for GraphicAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a GraphicAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

pub(crate) fn color_string(color: Rgb<u8>) -> String {
    format!("#{:02x}{:02x}{:02x}", color.r, color.g, color.b)
}

pub(crate) fn shadow_string(
    x_offset: Length,
    y_offset: Length,
    blur: Option<Length>,
    color: Rgb<u8>,
) -> String {
    if let Some(blur) = blur {
        format!("{} {} {} {}", color_string(color), x_offset, y_offset, blur)
    } else {
        format!("{} {} {}", color_string(color), x_offset, y_offset)
    }
}

pub(crate) fn rel_width_string(value: f64) -> String {
    format!("{}*", value)
}

pub(crate) fn border_string(width: Length, border: Border, color: Rgb<u8>) -> String {
    format!(
        "{} {} #{:02x}{:02x}{:02x}",
        width, border, color.r, color.g, color.b
    )
}

pub(crate) fn percent_string(value: f64) -> String {
    format!("{}%", value)
}

pub(crate) fn border_line_width_string(inner: Length, space: Length, outer: Length) -> String {
    format!("{} {} {}", inner, space, outer)
}
