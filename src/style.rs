///
/// Defines the basic structures for table styling, PageLayout and Style
///

use string_cache::DefaultAtom;

use crate::{CellRef, StyleFor, StyleOrigin, StyleUse};
use crate::attrmap::{AttrFoBackgroundColor, AttrFoBorder, AttrFoBreak, AttrFoKeepTogether, AttrFoKeepWithNext, AttrFoMargin, AttrFoMinHeight, AttrFontDecl, AttrFoPadding, AttrMap, AttrMapIter, AttrMapType, AttrParagraph, AttrStyleDynamicSpacing, AttrStyleShadow, AttrStyleWritingMode, AttrSvgHeight, AttrTableCell, AttrTableCol, AttrTableRow, AttrText};
use crate::text::TextVec;

/// Page layout.
/// Contains all header and footer information.
///
/// ```
/// use spreadsheet_ods::{write_ods, WorkBook};
/// use spreadsheet_ods::text::TextVec;
/// use spreadsheet_ods::style::{HeaderFooter, PageLayout};
/// use color::Rgb;
/// use spreadsheet_ods::attrmap::{AttrFoBackgroundColor, AttrFoMinHeight, AttrFoMargin};
///
/// let mut wb = WorkBook::new();
///
/// let mut pl = PageLayout::default();
///
/// pl.set_background_color(Rgb::new(12, 129, 252));
///
/// pl.header_attr_mut().set_min_height("0.75cm");
/// pl.header_attr_mut().set_margin_left("0.15cm");
/// pl.header_attr_mut().set_margin_right("0.15cm");
/// pl.header_attr_mut().set_margin_bottom("0.15cm");
///
/// pl.header_mut().center_mut().text("middle ground");
/// pl.header_mut().left_mut().text("left wing");
/// pl.header_mut().right_mut().text("right wing");
///
/// wb.add_pagelayout(pl);
///
/// write_ods(&wb, "test_out/hf0.ods").unwrap();
/// ```
///
#[derive(Clone, Debug, Default)]
pub struct PageLayout {
    name: String,
    masterpage_name: String,

    attr: Option<AttrMapType>,

    header_attr: HeaderFooterAttr,
    header: HeaderFooter,
    header_left: HeaderFooter,

    footer_attr: HeaderFooterAttr,
    footer: HeaderFooter,
    footer_left: HeaderFooter,
}

impl AttrMap for PageLayout {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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
    pub fn default() -> Self {
        Self {
            name: "Mpm1".to_string(),
            masterpage_name: "Default".to_string(),
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
    pub fn report() -> Self {
        Self {
            name: "Mpm2".to_string(),
            masterpage_name: "Report".to_string(),
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
    pub fn set_masterpage_name(&mut self, name: String) {
        self.masterpage_name = name;
    }

    /// In the xml pagelayout is split in two pieces. Each has a name.
    pub fn masterpage_name(&self) -> &String {
        &self.masterpage_name
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
    attr: Option<AttrMapType>,
}

impl AttrMap for HeaderFooterAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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
/// Each is a CompositVec of parsed XML-tags.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooter {
    display: bool,

    region_left: TextVec,
    region_center: TextVec,
    region_right: TextVec,

    content: TextVec,
}

impl HeaderFooter {
    /// Create
    pub fn new() -> Self {
        Self {
            display: true,
            region_left: TextVec::new(),
            region_center: TextVec::new(),
            region_right: TextVec::new(),
            content: TextVec::new(),
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
    pub fn set_left(&mut self, txt: TextVec) {
        self.region_left = txt;
    }

    /// Left region.
    pub fn left(&self) -> &TextVec {
        &self.region_left
    }

    /// Left region.
    pub fn left_mut(&mut self) -> &mut TextVec {
        &mut self.region_left
    }

    /// Center region.
    pub fn set_center(&mut self, txt: TextVec) {
        self.region_center = txt;
    }

    /// Center region.
    pub fn center(&self) -> &TextVec {
        &self.region_center
    }

    /// Center region.
    pub fn center_mut(&mut self) -> &mut TextVec {
        &mut self.region_center
    }

    /// Right region.
    pub fn set_right(&mut self, txt: TextVec) {
        self.region_right = txt;
    }

    /// Right region.
    pub fn right(&self) -> &TextVec {
        &self.region_right
    }

    /// Right region.
    pub fn right_mut(&mut self) -> &mut TextVec {
        &mut self.region_right
    }

    /// Header content, if there are no regions.
    pub fn set_content(&mut self, txt: TextVec) {
        self.content = txt;
    }

    /// Header content, if there are no regions.
    pub fn content(&self) -> &TextVec {
        &self.content
    }

    /// Header content, if there are no regions.
    pub fn content_mut(&mut self) -> &mut TextVec {
        &mut self.content
    }
}

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontFaceDecl {
    name: String,
    /// From where did we get this style.
    origin: StyleOrigin,
    /// All other attributes.
    attr: Option<AttrMapType>,
}

impl AttrFontDecl for FontFaceDecl {}

impl AttrMap for FontFaceDecl {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
        &mut self.attr
    }
}

impl FontFaceDecl {
    /// New, empty.
    pub fn new() -> Self {
        FontFaceDecl::new_origin(StyleOrigin::Content)
    }

    /// New, with origination.
    pub fn new_origin(origin: StyleOrigin) -> Self {
        Self {
            name: "".to_string(),
            origin,
            attr: None,
        }
    }

    /// New, with a name.
    pub fn with_name<S: Into<String>>(name: S) -> Self {
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

/// Style data fashioned after the ODS spec.
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
    /// Table styling
    table_attr: TableAttr,
    /// Column styling
    table_col_attr: TableColAttr,
    /// Row styling
    table_row_attr: TableRowAttr,
    /// Cell styles
    table_cell_attr: TableCellAttr,
    /// Cell paragraph styles
    paragraph_attr: ParagraphAttr,
    /// Cell text styles
    text_attr: TextAttr,
    /// Style maps
    style_map: Option<Vec<StyleMap>>,
}

impl Style {
    /// New, empty.
    pub fn new() -> Self {
        Style::new_origin(Default::default(), Default::default())
    }

    /// New, with origination.
    pub fn new_origin(origin: StyleOrigin, styleuse: StyleUse) -> Self {
        Style {
            name: String::from(""),
            display_name: None,
            origin,
            styleuse,
            family: Default::default(),
            parent: None,
            value_format: None,
            table_attr: Default::default(),
            table_col_attr: Default::default(),
            table_row_attr: Default::default(),
            table_cell_attr: Default::default(),
            paragraph_attr: Default::default(),
            text_attr: Default::default(),
            style_map: None,
        }
    }

    /// Creates a new cell style.
    /// value_style references a ValueFormat.
    pub fn cell_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableCell, name, value_style)
    }

    /// Creates a new column style.
    /// value_style references a ValueFormat.
    pub fn col_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableColumn, name, value_style)
    }

    /// Creates a new row style.
    /// value_style references a ValueFormat.
    pub fn row_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableRow, name, value_style)
    }

    /// Creates a new table style.
    /// value_style references a ValueFormat.
    pub fn table_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::Table, name, value_style)
    }

    /// New, with name.
    /// value_style references a ValueFormat.
    pub fn with_name<S: Into<String>, T: Into<String>>(family: StyleFor, name: S, value_style: T) -> Self {
        Style {
            name: name.into(),
            display_name: None,
            origin: Default::default(),
            styleuse: Default::default(),
            family,
            parent: Some(String::from("Default")),
            value_format: Some(value_style.into()),
            table_attr: Default::default(),
            table_col_attr: Default::default(),
            table_row_attr: Default::default(),
            table_cell_attr: Default::default(),
            paragraph_attr: Default::default(),
            text_attr: Default::default(),
            style_map: None,
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

    /// Table style attributes.
    pub fn table(&self) -> &TableAttr {
        &self.table_attr
    }

    /// Table style attributes.
    pub fn table_mut(&mut self) -> &mut TableAttr {
        &mut self.table_attr
    }

    /// Table column style attributes.
    pub fn col(&self) -> &TableColAttr {
        &self.table_col_attr
    }

    /// Table column style attributes.
    pub fn col_mut(&mut self) -> &mut TableColAttr {
        &mut self.table_col_attr
    }

    /// Table-row style attributes.
    pub fn row(&self) -> &TableRowAttr {
        &self.table_row_attr
    }

    /// Table-row style attributes.
    pub fn row_mut(&mut self) -> &mut TableRowAttr {
        &mut self.table_row_attr
    }

    /// Table-cell style attributes.
    pub fn cell(&self) -> &TableCellAttr {
        &self.table_cell_attr
    }

    /// Table-cell style attributes.
    pub fn cell_mut(&mut self) -> &mut TableCellAttr {
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

    /// Text style attributes.
    pub fn text(&self) -> &TextAttr {
        &self.text_attr
    }

    /// Text style attributes.
    pub fn text_mut(&mut self) -> &mut TextAttr {
        &mut self.text_attr
    }

    /// Adds a stylemap.
    pub fn add_stylemap(&mut self, stylemap: StyleMap) {
        if self.style_map.is_none() {
            self.style_map = Some(Vec::new());
        }
        if let Some(style_map) = &mut self.style_map {
            style_map.push(stylemap);
        }
    }

    /// Returns the stylemaps
    pub fn stylemaps(&self) -> Option<&Vec<StyleMap>> {
        self.style_map.as_ref()
    }
}

/// One style mapping.
/// The rules for this are not very clear. It writes the necessary data fine,
/// but the interpretation bei LO is not very accessible.
/// * The cellref must include a table-name.
/// * ???
/// * LO always adds calcext:conditional-formats which I can't handle.
///   I didn't find a spec for that.
/// TODO: clarify all of this.
#[derive(Clone, Debug, Default)]
pub struct StyleMap {
    condition: String,
    applied_style: String,
    base_cell: CellRef,
}

impl StyleMap {
    pub fn new<S: Into<String>, T: Into<String>>(condition: S, apply_style: T, cellref: CellRef) -> Self {
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


#[derive(Clone, Debug, Default)]
pub struct TableAttr {
    attr: Option<AttrMapType>,
}

impl AttrMap for TableAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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

#[derive(Clone, Debug, Default)]
pub struct TableRowAttr {
    attr: Option<AttrMapType>,
}

impl AttrMap for TableRowAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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

#[derive(Clone, Debug, Default)]
pub struct TableColAttr {
    attr: Option<AttrMapType>,
}

impl AttrMap for TableColAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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

#[derive(Clone, Debug, Default)]
pub struct TableCellAttr {
    attr: Option<AttrMapType>,
}

impl AttrMap for TableCellAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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

#[derive(Clone, Debug, Default)]
pub struct ParagraphAttr {
    attr: Option<AttrMapType>,
}

impl AttrMap for ParagraphAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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


#[derive(Clone, Debug, Default)]
pub struct TextAttr {
    attr: Option<AttrMapType>,
}

impl AttrMap for TextAttr {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
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
