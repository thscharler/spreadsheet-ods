use crate::style::pagestyle::PageStyleRef;
use crate::text::TextTag;

style_ref!(MasterPageRef);

/// Defines the structure and content for a page.
/// Refers to a PageStyle for layout information.
/// It must be attached to a Sheet to be used.
///
/// ```
/// use spreadsheet_ods::{pt, Length, WorkBook, Sheet};
/// use spreadsheet_ods::style::{PageStyle, MasterPage, TableStyle};
/// use spreadsheet_ods::style::units::Border;
/// use color::Rgb;
///
/// let mut wb = WorkBook::new();
///
/// let mut ps = PageStyle::new("ps1");
/// ps.set_border(pt!(0.5), Border::Groove, Rgb::new(128,128,128));
/// ps.headerstyle_mut().set_background_color(Rgb::new(92,92,92));
/// let ps_ref = wb.add_pagestyle(ps);
///
/// let mut mp1 = MasterPage::new("mp1");
/// mp1.set_pagestyle(&ps_ref);
/// mp1.header_mut().center_mut().add_text("center");
/// mp1.footer_mut().right_mut().add_text("right");
/// let mp1_ref = wb.add_masterpage(mp1);
///
/// let mut ts = TableStyle::new("ts1");
/// ts.set_master_page_name(&mp1_ref);
/// let ts_ref = wb.add_tablestyle(ts);
///
/// let mut sheet = Sheet::new();
/// sheet.set_style(&ts_ref);
/// ```  
///
#[derive(Clone, Debug, Default)]
pub struct MasterPage {
    name: String,
    pagestyle: String,

    header: HeaderFooter,
    header_first: HeaderFooter,
    header_left: HeaderFooter,

    footer: HeaderFooter,
    footer_first: HeaderFooter,
    footer_left: HeaderFooter,
}

impl MasterPage {
    /// Empty.
    pub fn empty() -> Self {
        Self {
            name: "".to_string(),
            pagestyle: "".to_string(),
            header: Default::default(),
            header_first: Default::default(),
            header_left: Default::default(),
            footer: Default::default(),
            footer_first: Default::default(),
            footer_left: Default::default(),
        }
    }

    /// New MasterPage
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            pagestyle: "".to_string(),
            header: Default::default(),
            header_first: Default::default(),
            header_left: Default::default(),
            footer: Default::default(),
            footer_first: Default::default(),
            footer_left: Default::default(),
        }
    }

    /// Style reference.
    pub fn masterpage_ref(&self) -> MasterPageRef {
        MasterPageRef::from(self.name())
    }

    /// Name.
    pub fn set_name(&mut self, name: String) {
        self.name = name;
    }

    /// Name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Reference to a page-style.
    pub fn set_pagestyle(&mut self, name: &PageStyleRef) {
        self.pagestyle = name.to_string();
    }

    /// Reference to a page-style.
    pub fn pagestyle(&self) -> &String {
        &self.pagestyle
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

    /// First page header.
    pub fn set_header_first(&mut self, header: HeaderFooter) {
        self.header_first = header;
    }

    /// First page header.
    pub fn header_first(&self) -> &HeaderFooter {
        &self.header_first
    }

    /// First page header.
    pub fn header_first_mut(&mut self) -> &mut HeaderFooter {
        &mut self.header_first
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

    /// First page footer.
    pub fn set_footer_first(&mut self, footer: HeaderFooter) {
        self.footer_first = footer;
    }

    /// First page footer.
    pub fn footer_first(&self) -> &HeaderFooter {
        &self.footer_first
    }

    /// First page footer.
    pub fn footer_first_mut(&mut self) -> &mut HeaderFooter {
        &mut self.footer_first
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

    /// true if all regions of the header/footer are empty.
    pub fn is_empty(&self) -> bool {
        self.region_left.is_none()
            && self.region_center.is_none()
            && self.region_right.is_none()
            && self.content.is_none()
    }

    /// Left region.
    pub fn set_left(&mut self, txt: TextTag) {
        self.region_left = Some(Box::new(txt));
    }

    /// Left region.
    pub fn left(&self) -> Option<&TextTag> {
        self.region_left.as_ref().map(|v| v.as_ref())
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
        self.region_center.as_ref().map(|v| v.as_ref())
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
        self.region_right.as_ref().map(|v| v.as_ref())
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
        self.content.as_ref().map(|v| v.as_ref())
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
