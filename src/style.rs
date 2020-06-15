///
/// Defines the basic structures for table styling, PageLayout and Style
///

use string_cache::DefaultAtom;

use crate::{StyleFor, StyleOrigin, StyleUse};
use crate::attrmap::{AttrFoBackgroundColor, AttrFoBorder, AttrFoBreak, AttrFoKeepTogether, AttrFoKeepWithNext, AttrFoMargin, AttrFoMinHeight, AttrFontDecl, AttrFoPadding, AttrMap, AttrMapIter, AttrMapType, AttrParagraph, AttrStyleDynamicSpacing, AttrStyleShadow, AttrStyleWritingMode, AttrSvgHeight, AttrTableCell, AttrTableCol, AttrTableRow, AttrText};
use crate::text::TextVec;

/// Page layout.
/// Contains all header and footer information.
#[derive(Clone, Debug, Default)]
pub struct PageLayout {
    pub(crate) name: String,
    pub(crate) masterpage_name: String,

    pub(crate) attr: Option<AttrMapType>,

    pub(crate) header_attr: HeaderFooterAttr,
    pub(crate) header: HeaderFooter,
    pub(crate) header_left: HeaderFooter,

    pub(crate) footer_attr: HeaderFooterAttr,
    pub(crate) footer: HeaderFooter,
    pub(crate) footer_left: HeaderFooter,
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

    /// Header. This is the regular header for left and right side pages.
    pub fn set_header(&mut self, header: HeaderFooter) {
        self.header = header;
    }

    /// Header.
    pub fn header(&self) -> &HeaderFooter {
        &self.header
    }

    /// If there is a different header on left side pages, set this one too.
    pub fn set_header_left(&mut self, header: HeaderFooter) {
        self.header_left = header;
    }

    /// Left side header.
    pub fn header_left(&self) -> &HeaderFooter {
        &self.header_left
    }

    /// Attributes for header.
    pub fn header_attr(&mut self) -> &mut HeaderFooterAttr {
        &mut self.header_attr
    }

    /// Footer. This is the regular footer for left and right side pages.
    pub fn set_footer(&mut self, footer: HeaderFooter) {
        self.footer = footer;
    }

    /// Footer.
    pub fn footer(&self) -> &HeaderFooter {
        &self.footer
    }

    /// If there is a different footer on left side pages, set this one too.
    pub fn set_footer_left(&mut self, footer: HeaderFooter) {
        self.footer_left = footer;
    }

    /// Left side footer.
    pub fn footer_left(&self) -> &HeaderFooter {
        &self.footer_left
    }

    /// Attributes for footer.
    pub fn footer_attr(&mut self) -> &mut HeaderFooterAttr {
        &mut self.footer_attr
    }
}

#[derive(Clone, Debug, Default)]
pub struct HeaderFooterAttr {
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) display: bool,

    pub(crate) region_left: TextVec,
    pub(crate) region_center: TextVec,
    pub(crate) region_right: TextVec,

    pub(crate) content: TextVec,
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

    /// Sets the left region.
    pub fn set_region_left(&mut self, region: TextVec) {
        self.region_left = region;
    }

    /// Left region.
    pub fn region_left(&self) -> &TextVec {
        &self.region_left
    }

    /// Sets the center region.
    pub fn set_region_center(&mut self, region: TextVec) {
        self.region_center = region;
    }

    /// Center region.
    pub fn region_center(&self) -> &TextVec {
        &self.region_center
    }

    /// Sets the right region.
    pub fn set_region_right(&mut self, region: TextVec) {
        self.region_right = region;
    }

    /// Right region.
    pub fn region_right(&self) -> &TextVec {
        &self.region_right
    }

    /// Sets the content for the whole header.
    pub fn set_content(&mut self, content: TextVec) {
        self.content = content;
    }

    /// Header content, if there are no regions.
    pub fn content(&mut self) -> &TextVec {
        &self.content
    }
}

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontFaceDecl {
    pub(crate) name: String,
    /// From where did we get this style.
    pub(crate) origin: StyleOrigin,
    /// All other attributes.
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) name: String,
    /// Nice String.
    pub(crate) display_name: Option<String>,
    /// From where did we get this style.
    pub(crate) origin: StyleOrigin,
    /// Which tag contains this style.
    pub(crate) styleuse: StyleUse,
    /// Applicability of this style.
    pub(crate) family: StyleFor,
    /// Styles can cascade.
    pub(crate) parent: Option<String>,
    /// References the actual formatting instructions in the value-styles.
    pub(crate) value_format: Option<String>,
    /// Table styling
    pub(crate) table_attr: TableAttr,
    /// Column styling
    pub(crate) table_col_attr: TableColAttr,
    /// Row styling
    pub(crate) table_row_attr: TableRowAttr,
    /// Cell styles
    pub(crate) table_cell_attr: TableCellAttr,
    /// Cell paragraph styles
    pub(crate) paragraph_attr: ParagraphAttr,
    /// Cell text styles
    pub(crate) text_attr: TextAttr,
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
    pub fn origin(&self) -> &StyleOrigin {
        &self.origin
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
    pub fn family(&self) -> &StyleFor {
        &self.family
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
    pub fn table_attr(&mut self) -> &mut TableAttr {
        &mut self.table_attr
    }

    pub fn col_attr(&mut self) -> &mut TableColAttr {
        &mut self.table_col_attr
    }

    /// Table-row style attributes.
    pub fn row_attr(&mut self) -> &mut TableRowAttr {
        &mut self.table_row_attr
    }

    /// Table-cell style attributes.
    pub fn cell_attr(&mut self) -> &mut TableCellAttr {
        &mut self.table_cell_attr
    }

    /// Paragraph style attributes.
    pub fn paragraph_attr(&mut self) -> &mut ParagraphAttr {
        &mut self.paragraph_attr
    }

    /// Text style attributes.
    pub fn text_attr(&mut self) -> &mut TextAttr {
        &mut self.text_attr
    }
}

#[derive(Clone, Debug, Default)]
pub struct TableAttr {
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) attr: Option<AttrMapType>,
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
    pub(crate) attr: Option<AttrMapType>,
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
