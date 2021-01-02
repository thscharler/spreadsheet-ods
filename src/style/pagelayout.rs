use crate::attrmap::{AttrMap, AttrMapIter, AttrMapType};
use crate::sealed::Sealed;
use crate::style::attr::{
    AttrFoBackgroundColor, AttrFoBorder, AttrFoMargin, AttrFoMinHeight, AttrFoPadding,
    AttrStyleDynamicSpacing, AttrStyleShadow, AttrSvgHeight,
};
use crate::text::TextTag;
use string_cache::DefaultAtom;

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

impl<'a> IntoIterator for &'a HeaderFooterAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
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
