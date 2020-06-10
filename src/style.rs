use std::collections::HashMap;
use std::fmt::{Display, Formatter};

use color::Rgb;
use string_cache::DefaultAtom;

use crate::{CompositVec, Sheet, StyleFor, StyleOrigin, StyleUse, ucell, WorkBook};
use crate::attrmap::{AttrFoBackground, AttrFoBorder, AttrFoMargin, AttrFoMinHeight, AttrFoPadding, AttrMap, AttrMapIter, AttrMapType, Border};
use crate::util::{clear_prp, get_prp, set_prp};

type PrpMap = HashMap<DefaultAtom, String>;

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

impl AttrFoBackground for PageLayout {}

impl AttrFoBorder for PageLayout {}

impl AttrFoMargin for PageLayout {}

impl AttrFoPadding for PageLayout {}


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
        self.masterpage_name = name.to_string();
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

impl AttrFoBackground for HeaderFooterAttr {}

impl AttrFoBorder for HeaderFooterAttr {}

impl AttrFoMargin for HeaderFooterAttr {}

impl AttrFoMinHeight for HeaderFooterAttr {}

impl AttrFoPadding for HeaderFooterAttr {}

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

    pub(crate) region_left: CompositVec,
    pub(crate) region_center: CompositVec,
    pub(crate) region_right: CompositVec,

    pub(crate) content: CompositVec,
}

impl HeaderFooter {
    /// Create
    pub fn new() -> Self {
        Self {
            display: true,
            region_left: CompositVec::new(),
            region_center: CompositVec::new(),
            region_right: CompositVec::new(),
            content: CompositVec::new(),
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
    pub fn set_region_left(&mut self, region: CompositVec) {
        self.region_left = region;
    }

    /// Left region.
    pub fn region_left(&self) -> &CompositVec {
        &self.region_left
    }

    /// Sets the center region.
    pub fn set_region_center(&mut self, region: CompositVec) {
        self.region_center = region;
    }

    /// Center region.
    pub fn region_center(&self) -> &CompositVec {
        &self.region_center
    }

    /// Sets the right region.
    pub fn set_region_right(&mut self, region: CompositVec) {
        self.region_right = region;
    }

    /// Right region.
    pub fn region_right(&self) -> &CompositVec {
        &self.region_right
    }

    /// Sets the content for the whole header.
    pub fn set_content(&mut self, content: CompositVec) {
        self.content = content;
    }

    /// Header content, if there are no regions.
    pub fn content(&mut self) -> &CompositVec {
        &self.content
    }
}

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontDecl {
    pub(crate) name: String,
    /// From where did we get this style.
    pub(crate) origin: StyleOrigin,
    /// All other attributes.
    pub(crate) prp: Option<HashMap<DefaultAtom, String>>,
}

impl FontDecl {
    /// New, empty.
    pub fn new() -> Self {
        FontDecl::new_origin(StyleOrigin::Content)
    }

    /// New, with origination.
    pub fn new_origin(origin: StyleOrigin) -> Self {
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
            origin: StyleOrigin::Content,
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

    /// Origin of the style
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Origin of the style
    pub fn origin(&self) -> StyleOrigin {
        self.origin
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
    pub(crate) table_prp: Option<PrpMap>,
    /// Column styling
    pub(crate) table_col_prp: Option<PrpMap>,
    /// Row styling
    pub(crate) table_row_prp: Option<PrpMap>,
    /// Cell styles
    pub(crate) table_cell_prp: Option<PrpMap>,
    /// Cell paragraph styles
    pub(crate) paragraph_prp: Option<PrpMap>,
    /// Cell text styles
    pub(crate) text_prp: Option<PrpMap>,
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
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
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
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
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

/// Sets the fontname for the font decl.
pub fn font_decl<S: Into<String>>(fontdecl: &mut FontDecl, family: S) {
    fontdecl.set_prp("svg:font-family", family.into());
}

/// Fontname for this style.
pub fn font_name<S: Into<String>>(style: &mut Style, font: S) {
    style.set_text_prp("style:font-name", font.into());
}

/// Font attributes for this style.
pub fn font_style(style: &mut Style, pt_size: f32, bold: bool, italic: bool) {
    font_size(style, pt_size);
    font_bold(style, bold);
    font_italic(style, italic);
}

/// Italic style.
pub fn font_italic(style: &mut Style, italic: bool) {
    if italic {
        style.set_text_prp("fo:font-style", "italic".to_string());
        style.set_text_prp("fo:font-style-asian", "italic".to_string());
        style.set_text_prp("fo:font-style-complex", "italic".to_string());
    } else {
        style.clear_text_prp("fo:font-style");
        style.clear_text_prp("fo:font-style-asian");
        style.clear_text_prp("fo:font-style-complex");
    }
}

/// Bold style.
pub fn font_bold(style: &mut Style, bold: bool) {
    if bold {
        style.set_text_prp("fo:font-weight", "bold".to_string());
        style.set_text_prp("fo:font-weight-asian", "bold".to_string());
        style.set_text_prp("fo:font-weight-complex", "bold".to_string());
    } else {
        style.clear_text_prp("fo:font-weight");
        style.clear_text_prp("fo:font-weight-asian");
        style.clear_text_prp("fo:font-weight-complex");
    }
}

/// Font size.
pub fn font_size(style: &mut Style, pt_size: f32) {
    style.set_text_prp("fo:font-size", format!("{}pt", pt_size));
    style.set_text_prp("fo:font-size-asian", format!("{}pt", pt_size));
    style.set_text_prp("fo:font-size-complex", format!("{}pt", pt_size));
}

/// Font color.
pub fn font_color(style: &mut Style, color: Rgb<u8>) {
    style.set_text_prp("fo:color", color_string(color));
}

/// Various underline styles.
#[derive(Debug, Clone, Copy)]
pub enum Underline {
    Solid,
    Double,
    Dotted,
    Dashed,
    Wavy,
}

impl Display for Underline {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Underline::Solid => write!(f, "solid"),
            Underline::Double => write!(f, "double"),
            Underline::Dotted => write!(f, "dotted"),
            Underline::Dashed => write!(f, "dashed"),
            Underline::Wavy => write!(f, "wavy"),
        }
    }
}

/// Underline style.
pub fn font_underline(style: &mut Style, ustyle: Underline) {
    style.set_text_prp("style:text-underline-style", ustyle.to_string());
    style.set_text_prp("style:text-underline-width", "auto".to_string());
    style.set_text_prp("style:text-underline-color", "font-color".to_string());
}

/// Various line-throug styles.
#[derive(Debug, Clone, Copy)]
pub enum LineThroughStyle {
    Dashed,
    DotDash,
    DotDotDash,
    Dotted,
    LongDash,
    None,
    Solid,
    Wave,
}

impl Display for LineThroughStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineThroughStyle::Dashed => write!(f, "dashed"),
            LineThroughStyle::DotDash => write!(f, "dot-dash"),
            LineThroughStyle::DotDotDash => write!(f, "dot-dot-dash"),
            LineThroughStyle::Dotted => write!(f, "dotted"),
            LineThroughStyle::LongDash => write!(f, "long-dash"),
            LineThroughStyle::None => write!(f, "none"),
            LineThroughStyle::Solid => write!(f, "solid"),
            LineThroughStyle::Wave => write!(f, "wavae"),
        }
    }
}

/// Various line-through types.
#[derive(Debug, Clone, Copy)]
pub enum LineThroughType {
    None,
    Single,
    Double,
}

impl Display for LineThroughType {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineThroughType::None => write!(f, "none"),
            LineThroughType::Single => write!(f, "single"),
            LineThroughType::Double => write!(f, "double"),
        }
    }
}

/// Sets the line-through style
pub fn font_line_through(style: &mut Style, ltstyle: LineThroughStyle, lttype: LineThroughType) {
    style.set_text_prp("style:text-line-through-style", ltstyle.to_string());
    style.set_text_prp("style:text-line-through-type", lttype.to_string());
}

/// Font as outline.
pub fn font_outline(style: &mut Style, outline: bool) {
    style.set_text_prp("style:text-outline", outline.to_string());
}

/// Font with shadow.
pub fn font_shadow(style: &mut Style, pt_shadow_x: f32, pt_shadow_y: f32) {
    style.set_text_prp("fo:text-shadow", format!("{}pt {}pt", pt_shadow_x, pt_shadow_y));
}

// format as string
fn border_string(width: f32, border: Border, color: Rgb<u8>) -> String {
    format!("{}pt {} #{:02x}{:02x}{:02x}", width, border, color.r, color.g, color.b)
}

// format as string
fn color_string(color: Rgb<u8>) -> String {
    format!(" #{:02x}{:02x}{:02x}", color.r, color.g, color.b)
}

/// Border style all four sides.
pub fn cell_border(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border", border_string(pt_width, border, color));
}

/// Border style.
pub fn cell_border_bottom(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-bottom", border_string(pt_width, border, color));
}

/// Border style.
pub fn cell_border_top(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-top", border_string(pt_width, border, color));
}

/// Border style.
pub fn cell_border_left(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-left", border_string(pt_width, border, color));
}

/// Border style.
pub fn cell_border_right(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-right", border_string(pt_width, border, color));
}

/// Border style.
pub fn cell_background(style: &mut Style, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:background-color", color_string(color));
}

/// Border style.
pub fn cell_padding(style: &mut Style, pt_padding: f32) {
    style.set_table_cell_prp("fo:padding", format!("{}pt", pt_padding));
}

/// Border style.
pub fn cell_shadow(style: &mut Style, pt_off_x: f32, pt_off_y: f32, color: Rgb<u8>) {
    style.set_table_cell_prp("style:shadow", format!("#{:02x}{:02x}{:02x} {}pt {}pt", color.r, color.g, color.b, pt_off_x, pt_off_y));
}

/// Border style.
pub fn cell_shrink_to_fit(style: &mut Style, shrink: bool) {
    style.set_table_cell_prp("style:shrink-to-fit", shrink.to_string());
}

/// Border style.
pub fn cell_rotation_angle(style: &mut Style, angle: f32) {
    style.set_table_cell_prp("style:rotation-angle", angle.to_string());
}

/// Horizontal alignment.
pub enum Align {
    Start,
    Center,
    End,
    Justify,
    Inside,
    Outside,
    Left,
    Right,
}

impl Display for Align {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Align::Start => write!(f, "start"),
            Align::Center => write!(f, "center"),
            Align::End => write!(f, "end"),
            Align::Justify => write!(f, "justify"),
            Align::Inside => write!(f, "inside"),
            Align::Outside => write!(f, "outside"),
            Align::Left => write!(f, "left"),
            Align::Right => write!(f, "right"),
        }
    }
}

/// Alignment
pub fn text_align(style: &mut Style, align: Align) {
    style.set_paragraph_prp("fo:text-align", align.to_string());
}

/// Column width.
pub fn set_col_width(workbook: &mut WorkBook, sheet: &mut Sheet, col: ucell, width: &str) {
    let style_name = format!("co{}", col);

    let mut col_style = Style::with_name(StyleFor::TableColumn, &style_name, "");
    col_style.set_table_col_prp("style:column-width", width.to_string());
    workbook.add_style(col_style);

    sheet.set_column_style(col, &style_name);
}

/// Row height.
pub fn set_row_height(workbook: &mut WorkBook, sheet: &mut Sheet, row: ucell, height: &str) {
    let style_name = format!("ro{}", row);

    let mut row_style = Style::row_style(&style_name, "");
    row_style.set_table_row_prp("style:row-height", height.to_string());
    row_style.set_table_row_prp("style:use-optimal-row-height", "false".to_string());
    workbook.add_style(row_style);

    sheet.set_row_style(row, &style_name);
}