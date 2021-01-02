use crate::attrmap2::AttrMap2;
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string, Border,
    Length,
};
use crate::text::TextTag;
use color::Rgb;

/// Page layout.
/// Contains all header and footer information.
///
/// ```
/// use spreadsheet_ods::{write_ods, WorkBook};
/// use spreadsheet_ods::{cm};
/// use spreadsheet_ods::style::{HeaderFooter, PageLayout};
/// use color::Rgb;
/// use spreadsheet_ods::style::units::Length;
///
/// let mut wb = WorkBook::new();
///
/// let mut pl = PageLayout::new_default();
///
/// pl.set_background_color(Rgb::new(12, 129, 252));
///
/// pl.header_style_mut().set_min_height(cm!(0.75));
/// pl.header_style_mut().set_margin_left(cm!(0.15));
/// pl.header_style_mut().set_margin_right(cm!(0.15));
/// pl.header_style_mut().set_margin_bottom(Length::Cm(0.75));
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

    style: AttrMap2,

    header_style: HeaderFooterStyle,
    header: HeaderFooter,
    header_left: HeaderFooter,

    footer_style: HeaderFooterStyle,
    footer: HeaderFooter,
    footer_left: HeaderFooter,
}

impl PageLayout {
    /// Create with name "Mpm1" and masterpage-name "Default".
    pub fn new_default() -> Self {
        Self {
            name: "Mpm1".to_string(),
            master_page_name: "Default".to_string(),
            style: Default::default(),
            header: Default::default(),
            header_left: Default::default(),
            header_style: Default::default(),
            footer: Default::default(),
            footer_left: Default::default(),
            footer_style: Default::default(),
        }
    }

    /// Create with name "Mpm2" and masterpage-name "Report".
    pub fn new_report() -> Self {
        Self {
            name: "Mpm2".to_string(),
            master_page_name: "Report".to_string(),
            style: Default::default(),
            header: Default::default(),
            header_left: Default::default(),
            header_style: Default::default(),
            footer: Default::default(),
            footer_left: Default::default(),
            footer_style: Default::default(),
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

    pub fn style(&self) -> &AttrMap2 {
        &self.style
    }

    pub fn style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.style
    }
    // TODO: more attributes
    // fo:page-height 20.215, fo:page-width 20.216, style:border-line-width 20.248,
    // style:border-line-width-bottom 20.249, style:border-line-width-left 20.250,
    // style:border-line-width-right 20.251, style:border-line-width-top 20.252,
    // style:first-page-number 20.266, style:footnote-max-height 20.296,
    // style:layout-grid-base-height 20.304, style:layout-grid-base-width 20.305,
    // style:layout-grid-color 20.306, style:layout-grid-display 20.307,
    // style:layout-grid-lines 20.308, style:layout-grid-mode 20.309,
    // style:layoutgrid-print 20.310, style:layout-grid-ruby-below 20.311,
    // style:layout-gridruby-height 20.312, style:layout-grid-snap-to 20.313,
    // style:layout-gridstandard-mode 20.314,
    // style:num-format 20.322, style:num-letter-sync 20.323,
    // style:num-prefix 20.324, style:num-suffix 20.325, style:paper-tray-name
    // 20.329, style:print 20.330, style:print-orientation 20.333, style:print-pageorder 20.332,
    // style:register-truth-ref-style-name 20.337, style:scale-to
    // 20.352, style:scale-to-X 20.354, style:scale-to-Y 20.355, style:scale-to-pages
    // 20.353, style:shadow 20.359, style:table-centering 20.363 and style:writingmode 20.404.
    fo_background_color!(style_mut);
    fo_border!(style_mut);
    fo_margin!(style_mut);
    fo_padding!(style_mut);
    style_dynamic_spacing!(style_mut);
    style_shadow!(style_mut);
    svg_height!(style_mut);

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
    pub fn header_style(&self) -> &HeaderFooterStyle {
        &self.header_style
    }

    /// Attributes for header.
    pub fn header_style_mut(&mut self) -> &mut HeaderFooterStyle {
        &mut self.header_style
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
    pub fn footer_style(&self) -> &HeaderFooterStyle {
        &self.footer_style
    }

    /// Attributes for footer.
    pub fn footer_style_mut(&mut self) -> &mut HeaderFooterStyle {
        &mut self.footer_style
    }
}

#[derive(Clone, Debug, Default)]
pub struct HeaderFooterStyle {
    style: AttrMap2,
}

impl HeaderFooterStyle {
    pub fn style(&self) -> &AttrMap2 {
        &self.style
    }

    pub fn style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.style
    }

    // TODO: more of this ...
    // style:border-line-width 20.248,
    // style:border-linewidth-bottom 20.249, style:border-line-width-left 20.250,
    // style:border-linewidth-right 20.251, style:border-line-width-top 20.252,
    fo_background_color!(style_mut);
    fo_border!(style_mut);
    fo_margin!(style_mut);
    fo_min_height!(style_mut);
    fo_padding!(style_mut);
    style_dynamic_spacing!(style_mut);
    style_shadow!(style_mut);
    svg_height!(style_mut);
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
