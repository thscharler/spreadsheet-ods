//!
//! Defines the basic structures for table styling,
//! [PageLayout](struct.PageLayout.html)
//! and [Style](struct.Style.html)
//!

use string_cache::DefaultAtom;

use crate::text::TextTag;
use crate::CellRef;

pub use crate::attrmap::*;
use crate::sealed::Sealed;
use crate::style::color_string;
use color::Rgb;
use std::fmt::{Display, Formatter};

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
    None,
}

impl Default for StyleFor {
    fn default() -> Self {
        StyleFor::None
    }
}

/// Page layout.
/// Contains all header and footer information.
///
/// ```
/// use spreadsheet_ods::{write_ods, WorkBook};
/// use spreadsheet_ods::{cm};
/// use spreadsheet_ods::style::{HeaderFooter, PageLayout, Length};
/// use color::Rgb;
/// use spreadsheet_ods::style::{AttrFoBackgroundColor, AttrFoMinHeight, AttrFoMargin};
///
/// let mut wb = WorkBook::new();
///
/// let mut pl = PageLayout::new_default();
///
/// pl.set_background_color(Rgb::new(12, 129, 252));
///
/// pl.header_attr_mut().set_min_height(cm!(0.75));
/// pl.header_attr_mut().set_margin_left(cm!(0.15));
/// pl.header_attr_mut().set_margin_right(cm!(0.15));
/// pl.header_attr_mut().set_margin_bottom(Length::Cm(0.75));
///
/// pl.header_mut().center_mut().push_text("middle ground");
/// pl.header_mut().left_mut().push_text("left wing");
/// pl.header_mut().right_mut().push_text("right wing");
///
/// wb.add_pagelayout(pl);
///
/// write_ods(&wb, "test_out/hf0.ods").unwrap();
/// ```
///
#[derive(Clone, Debug, Default)]
pub struct PageLayout {
    name: String,
    master_page_name: String,

    attr: AttrMapType,

    header_attr: HeaderFooterAttr,
    header: HeaderFooter,
    header_left: HeaderFooter,

    footer_attr: HeaderFooterAttr,
    footer: HeaderFooter,
    footer_left: HeaderFooter,
}

impl Sealed for PageLayout {}

impl AttrMap for PageLayout {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl AttrFoBackgroundColor for PageLayout {}

impl AttrFoBorder for PageLayout {}

impl AttrFoMargin for PageLayout {}

impl AttrFoPadding for PageLayout {}

impl AttrStyleDynamicSpacing for PageLayout {}

impl AttrStyleShadow for PageLayout {}

impl AttrSvgHeight for PageLayout {}

impl PageLayout {
    /// Create with name "Mpm1" and masterpage-name "Default".
    pub fn new_default() -> Self {
        Self {
            name: "Mpm1".to_string(),
            master_page_name: "Default".to_string(),
            attr: None,
            header: Default::default(),
            header_left: Default::default(),
            header_attr: Default::default(),
            footer: Default::default(),
            footer_left: Default::default(),
            footer_attr: Default::default(),
        }
    }

    /// Create with name "Mpm2" and masterpage-name "Report".
    pub fn new_report() -> Self {
        Self {
            name: "Mpm2".to_string(),
            master_page_name: "Report".to_string(),
            attr: None,
            header: Default::default(),
            header_left: Default::default(),
            header_attr: Default::default(),
            footer: Default::default(),
            footer_left: Default::default(),
            footer_attr: Default::default(),
        }
    }

    /// Name.
    pub fn set_name(&mut self, name: String) {
        self.name = name;
    }

    /// Name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// In the xml pagelayout is split in two pieces. Each has a name.
    pub fn set_master_page_name(&mut self, name: String) {
        self.master_page_name = name;
    }

    /// In the xml pagelayout is split in two pieces. Each has a name.
    pub fn master_page_name(&self) -> &String {
        &self.master_page_name
    }

    /// Iterator over the attributes of this pagelayout.
    pub fn attr_iter(&self) -> AttrMapIter {
        AttrMapIter::from(self.attr_map())
    }

    /// Left side header.
    pub fn set_header(&mut self, header: HeaderFooter) {
        self.header = header;
    }

    /// Left side header.
    pub fn header(&self) -> &HeaderFooter {
        &self.header
    }

    /// Header.
    pub fn header_mut(&mut self) -> &mut HeaderFooter {
        &mut self.header
    }

    /// Left side header.
    pub fn set_header_left(&mut self, header: HeaderFooter) {
        self.header_left = header;
    }

    /// Left side header.
    pub fn header_left(&self) -> &HeaderFooter {
        &self.header_left
    }

    /// Left side header.
    pub fn header_left_mut(&mut self) -> &mut HeaderFooter {
        &mut self.header_left
    }

    /// Attributes for header.
    pub fn header_attr(&self) -> &HeaderFooterAttr {
        &self.header_attr
    }

    /// Attributes for header.
    pub fn header_attr_mut(&mut self) -> &mut HeaderFooterAttr {
        &mut self.header_attr
    }

    /// Footer.
    pub fn set_footer(&mut self, footer: HeaderFooter) {
        self.footer = footer;
    }

    /// Footer.
    pub fn footer(&self) -> &HeaderFooter {
        &self.footer
    }

    /// Footer.
    pub fn footer_mut(&mut self) -> &mut HeaderFooter {
        &mut self.footer
    }

    /// Left side footer.
    pub fn set_footer_left(&mut self, footer: HeaderFooter) {
        self.footer_left = footer;
    }

    /// Left side footer.
    pub fn footer_left(&self) -> &HeaderFooter {
        &self.footer_left
    }

    /// Left side footer.
    pub fn footer_left_mut(&mut self) -> &mut HeaderFooter {
        &mut self.footer_left
    }

    /// Attributes for footer.
    pub fn footer_attr(&self) -> &HeaderFooterAttr {
        &self.footer_attr
    }

    /// Attributes for footer.
    pub fn footer_attr_mut(&mut self) -> &mut HeaderFooterAttr {
        &mut self.footer_attr
    }
}

#[derive(Clone, Debug, Default)]
pub struct HeaderFooterAttr {
    attr: AttrMapType,
}

impl Sealed for HeaderFooterAttr {}

impl AttrMap for HeaderFooterAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl AttrFoBackgroundColor for HeaderFooterAttr {}

impl AttrFoBorder for HeaderFooterAttr {}

impl AttrFoMargin for HeaderFooterAttr {}

impl AttrFoMinHeight for HeaderFooterAttr {}

impl AttrFoPadding for HeaderFooterAttr {}

impl AttrStyleDynamicSpacing for HeaderFooterAttr {}

impl AttrStyleShadow for HeaderFooterAttr {}

impl AttrSvgHeight for HeaderFooterAttr {}

impl<'a> IntoIterator for &'a HeaderFooterAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

/// Header/Footer data.
/// Can be seen as three regions left/center/right or as one region.
/// In the first case region* contains the data, in the second it's content.
/// Each is a TextTag of parsed XML-tags.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooter {
    display: bool,

    region_left: Option<Box<TextTag>>,
    region_center: Option<Box<TextTag>>,
    region_right: Option<Box<TextTag>>,

    content: Option<Box<TextTag>>,
}

impl HeaderFooter {
    /// Create
    pub fn new() -> Self {
        Self {
            display: true,
            region_left: None,
            region_center: None,
            region_right: None,
            content: None,
        }
    }

    /// Is the header displayed. Used to deactivate left side headers.
    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }

    /// Display
    pub fn display(&self) -> bool {
        self.display
    }

    /// Left region.
    pub fn set_left(&mut self, txt: TextTag) {
        self.region_left = Some(Box::new(txt));
    }

    /// Left region.
    pub fn left(&self) -> Option<&TextTag> {
        match &self.region_left {
            None => None,
            Some(v) => Some(v.as_ref()),
        }
    }

    /// Left region.
    pub fn left_mut(&mut self) -> &mut TextTag {
        if self.region_left.is_none() {
            self.region_left = Some(Box::new(TextTag::new("text:p")));
        }
        if let Some(center) = &mut self.region_left {
            center
        } else {
            unreachable!()
        }
    }

    /// Center region.
    pub fn set_center(&mut self, txt: TextTag) {
        self.region_center = Some(Box::new(txt));
    }

    /// Center region.
    pub fn center(&self) -> Option<&TextTag> {
        match &self.region_center {
            None => None,
            Some(v) => Some(v.as_ref()),
        }
    }

    /// Center region.
    pub fn center_mut(&mut self) -> &mut TextTag {
        if self.region_center.is_none() {
            self.region_center = Some(Box::new(TextTag::new("text:p")));
        }
        if let Some(center) = &mut self.region_center {
            center
        } else {
            unreachable!()
        }
    }

    /// Right region.
    pub fn set_right(&mut self, txt: TextTag) {
        self.region_right = Some(Box::new(txt));
    }

    /// Right region.
    pub fn right(&self) -> Option<&TextTag> {
        match &self.region_right {
            None => None,
            Some(v) => Some(v.as_ref()),
        }
    }

    /// Right region.
    pub fn right_mut(&mut self) -> &mut TextTag {
        if self.region_right.is_none() {
            self.region_right = Some(Box::new(TextTag::new("text:p")));
        }
        if let Some(center) = &mut self.region_right {
            center
        } else {
            unreachable!()
        }
    }

    /// Header content, if there are no regions.
    pub fn set_content(&mut self, txt: TextTag) {
        self.content = Some(Box::new(txt));
    }

    /// Header content, if there are no regions.
    pub fn content(&self) -> Option<&TextTag> {
        match &self.content {
            None => None,
            Some(v) => Some(v.as_ref()),
        }
    }

    /// Header content, if there are no regions.
    pub fn content_mut(&mut self) -> &mut TextTag {
        if self.content.is_none() {
            self.content = Some(Box::new(TextTag::new("text:p")));
        }
        if let Some(center) = &mut self.content {
            center
        } else {
            unreachable!()
        }
    }
}

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontFaceDecl {
    name: String,
    /// From where did we get this style.
    origin: StyleOrigin,
    /// All other attributes.
    attr: AttrMapType,
}

impl Sealed for FontFaceDecl {}

impl AttrFontDecl for FontFaceDecl {}

impl AttrMap for FontFaceDecl {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl FontFaceDecl {
    /// New, empty.
    pub fn new() -> Self {
        Self {
            name: "".to_string(),
            origin: Default::default(),
            attr: None,
        }
    }

    /// New, with a name.
    pub fn new_with_name<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            origin: StyleOrigin::Content,
            attr: None,
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

    /// Origin of the style
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Origin of the style
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Iterator over the attributes of this pagelayout.
    pub fn attr_iter(&self) -> AttrMapIter {
        AttrMapIter::from(self.attr_map())
    }
}

/// Styles define a large number of attributes. These are grouped together
/// as table, row, column, cell, paragraph and text attributes.
///
/// ```
/// use spreadsheet_ods::{Style, CellRef, WorkBook};
/// use spreadsheet_ods::style::{StyleOrigin, StyleUse, AttrText, StyleMap};
/// use color::Rgb;
///
/// let mut wb = WorkBook::new();
///
/// let mut st = Style::new_cell_style("ce12", "num2");
/// st.text_mut().set_color(Rgb::new(192, 128, 0));
/// st.text_mut().set_font_bold();
/// wb.add_style(st);
///
/// let mut st = Style::new_cell_style("ce11", "num2");
/// st.text_mut().set_color(Rgb::new(0, 192, 128));
/// st.text_mut().set_font_bold();
/// wb.add_style(st);
///
/// let mut st = Style::new_cell_style("ce13", "num4");
/// st.push_stylemap(StyleMap::new("cell-content()=\"BB\"", "ce12", CellRef::remote("sheet0", 4, 3)));
/// st.push_stylemap(StyleMap::new("cell-content()=\"CC\"", "ce11", CellRef::remote("sheet0", 4, 3)));
/// wb.add_style(st);
/// ```
/// Styles can be defined in content.xml or as global styles in styles.xml. This
/// is reflected as the StyleOrigin. The StyleUse differentiates between automatic
/// and user visible, named styles. And third StyleFor defines for which part of
/// the document the style can be used.
///
/// Cell styles usually reference a value format for text formatting purposes.
///
/// Styles can also link to a parent style and to a pagelayout.
///
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
        assert_eq!(self.family, StyleFor::Table, "Can only be used for Table-Style.");
        &self.table_attr
    }

    /// Table style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::Table.    
    pub fn table_mut(&mut self) -> &mut TableAttr {
        assert_eq!(self.family, StyleFor::Table, "Can only be used for Table-Style.");
        &mut self.table_attr
    }

    /// Table column style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableColumn.
    pub fn col(&self) -> &TableColAttr {
        assert_eq!(self.family, StyleFor::TableColumn, "Can only be used for Column-Style.");
        &self.table_col_attr
    }

    /// Table column style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableColumn.
    pub fn col_mut(&mut self) -> &mut TableColAttr {
        assert_eq!(self.family, StyleFor::TableColumn, "Can only be used for Column-Style.");
        &mut self.table_col_attr
    }

    /// Table-row style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableRow.
    pub fn row(&self) -> &TableRowAttr {
        assert_eq!(self.family, StyleFor::TableRow, "Can only be used for Row-Style.");
        &self.table_row_attr
    }

    /// Table-row style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableRow.
    pub fn row_mut(&mut self) -> &mut TableRowAttr {
        assert_eq!(self.family, StyleFor::TableRow, "Can only be used for Row-Style.");
        &mut self.table_row_attr
    }

    /// Table-cell style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableCell.    
    pub fn cell(&self) -> &TableCellAttr {
        assert_eq!(self.family, StyleFor::TableCell, "Can only be used for Cell-Style.");
        &self.table_cell_attr
    }

    /// Table-cell style attributes.
    ///
    /// Panic
    ///
    /// Only accessible when family() == StyleFor::TableCell.    
    pub fn cell_mut(&mut self) -> &mut TableCellAttr {
        assert_eq!(self.family, StyleFor::TableCell, "Can only be used for Cell-Style.");
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

/// One style mapping.
///
/// The rules for this are not very clear. It writes the necessary data fine,
/// but the interpretation by LibreOffice is not very intelligable.
///
/// * The base-cell must include a table-name.
/// * LibreOffice always adds calcext:conditional-formats which I can't handle.
///
/// TODO: clarify all of this.
///
#[derive(Clone, Debug, Default)]
pub struct StyleMap {
    condition: String,
    applied_style: String,
    base_cell: CellRef,
}

impl StyleMap {
    pub fn new<S: Into<String>, T: Into<String>>(
        condition: S,
        apply_style: T,
        cellref: CellRef,
    ) -> Self {
        Self {
            condition: condition.into(),
            applied_style: apply_style.into(),
            base_cell: cellref,
        }
    }

    pub fn condition(&self) -> &String {
        &self.condition
    }

    pub fn set_condition<S: Into<String>>(&mut self, cond: S) {
        self.condition = cond.into();
    }

    pub fn applied_style(&self) -> &String {
        &self.applied_style
    }

    pub fn set_applied_style<S: Into<String>>(&mut self, style: S) {
        self.applied_style = style.into();
    }

    pub fn base_cell(&self) -> &CellRef {
        &self.base_cell
    }

    pub fn set_base_cell(&mut self, cellref: CellRef) {
        self.base_cell = cellref;
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

#[derive(Clone, Copy, Debug)]
pub enum TabStopType {
    Center,
    Left,
    Right,
    Char,
}

impl Display for TabStopType {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TabStopType::Center => write!(f, "center"),
            TabStopType::Left => write!(f, "left"),
            TabStopType::Right => write!(f, "right"),
            TabStopType::Char => write!(f, "char"),
        }
    }
}

impl Default for TabStopType {
    fn default() -> Self {
        Self::Left
    }
}

/// Tabstops are part of a paragraph style.
#[derive(Clone, Debug, Default)]
pub struct TabStop {
    attr: AttrMapType,
}

impl TabStop {
    pub fn new() -> Self {
        Self {
            attr: Default::default(),
        }
    }

    /// Delimiter character for tabs of type Char.
    pub fn set_tabstop_char(&mut self, c: char) {
        self.set_attr("style:char", c.to_string());
    }

    /// Color
    pub fn set_leader_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:leader-color", color_string(color));
    }

    /// Linestyle for the leader line.
    pub fn set_leader_style(&mut self, style: LineStyle) {
        self.set_attr("style:leader-style", style.to_string());
    }

    /// Fill character for the leader line.
    pub fn set_leader_text(&mut self, text: char) {
        self.set_attr("style:leader-text", text.to_string());
    }

    /// Textstyle for the leader line.
    pub fn set_leader_text_style(&mut self, styleref: String) {
        self.set_attr("style:leader-text-style", styleref);
    }

    /// LineType for the leader line.
    pub fn set_leader_type(&mut self, t: LineType) {
        self.set_attr("style:leader-type", t.to_string());
    }

    /// Width of the leader line.
    pub fn set_leader_width(&mut self, w: LineWidth) {
        self.set_attr("style:leader-width", w.to_string());
    }

    /// Position of the tab stop.
    pub fn set_position(&mut self, pos: Length) {
        self.set_attr("style:position", pos.to_string());
    }

    /// Type of the tab stop.
    pub fn set_tabstop_type(&mut self, t: TabStopType) {
        self.set_attr("style:type", t.to_string());
    }
}

impl Sealed for TabStop {}

impl AttrMap for TabStop {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

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
