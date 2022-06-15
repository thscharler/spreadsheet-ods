use crate::attrmap2::AttrMap2;
use crate::style::units::{Border, PrintOrientation};
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string,
};
use crate::Length;
use color::Rgb;

style_ref!(PageStyleRef);

/// Describes the style information for a page.
/// For an example see MasterPage.
///
#[derive(Debug, Clone)]
pub struct PageStyle {
    name: String,
    style: AttrMap2,
    header: HeaderFooterStyle,
    footer: HeaderFooterStyle,
}

impl PageStyle {
    /// New pagestyle.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            style: Default::default(),
            header: Default::default(),
            footer: Default::default(),
        }
    }

    /// Style reference.
    pub fn style_ref(&self) -> PageStyleRef {
        PageStyleRef::from(self.name())
    }

    /// Style name
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Style name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Attributes for header.
    pub fn headerstyle(&self) -> &HeaderFooterStyle {
        &self.header
    }

    /// Attributes for header.
    pub fn headerstyle_mut(&mut self) -> &mut HeaderFooterStyle {
        &mut self.header
    }

    /// Attributes for footer.
    pub fn footerstyle(&self) -> &HeaderFooterStyle {
        &self.footer
    }

    /// Attributes for footer.
    pub fn footerstyle_mut(&mut self) -> &mut HeaderFooterStyle {
        &mut self.footer
    }

    /// Access to all style attributes.
    pub fn style(&self) -> &AttrMap2 {
        &self.style
    }

    /// Access to all style attributes.
    pub fn style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.style
    }

    // style:first-page-number 20.266,
    // style:footnote-max-height 20.296,
    // style:layout-grid-base-height 20.304,
    // style:layout-grid-base-width 20.305,
    // style:layout-grid-color 20.306,
    // style:layout-grid-display 20.307,
    // style:layout-grid-lines 20.308,
    // style:layout-grid-mode 20.309,
    // style:layoutgrid-print 20.310,
    // style:layout-grid-ruby-below 20.311,
    // style:layout-gridruby-height 20.312,
    // style:layout-grid-snap-to 20.313,
    // style:layout-gridstandard-mode 20.314,
    // style:num-format 20.322,
    // style:num-letter-sync 20.323,
    // style:num-prefix 20.324,
    // style:num-suffix 20.325,
    // style:paper-tray-name 20.329,
    // style:print 20.330,
    // ok style:print-orientation 20.333,
    // style:print-pageorder 20.332,
    // style:register-truth-ref-style-name 20.337,
    // style:scale-to 20.352,
    // style:scale-to-X 20.354,
    // style:scale-to-Y 20.355,
    // style:scale-to-pages 20.353,
    // style:shadow 20.359,
    // style:table-centering 20.363
    // style:writingmode 20.404.

    /// Print orientation
    pub fn set_print_orientation(&mut self, orientation: PrintOrientation) {
        self.style_mut()
            .set_attr("style:print-orientation", orientation.to_string());
    }

    /// Page Height
    pub fn set_page_height(&mut self, height: Length) {
        self.style_mut()
            .set_attr("fo:page-height", height.to_string());
    }

    /// Page Width
    pub fn set_page_width(&mut self, width: Length) {
        self.style_mut()
            .set_attr("fo:page-width", width.to_string());
    }

    fo_background_color!(style_mut);
    fo_border!(style_mut);
    fo_margin!(style_mut);
    fo_padding!(style_mut);
    style_dynamic_spacing!(style_mut);
    style_shadow!(style_mut);
    svg_height!(style_mut);
}

/// Style attributes for header/footer.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooterStyle {
    style: AttrMap2,
}

impl HeaderFooterStyle {
    /// General attributes.
    pub fn style(&self) -> &AttrMap2 {
        &self.style
    }

    /// General attributes.
    pub fn style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.style
    }

    fo_background_color!(style_mut);
    fo_border!(style_mut);
    fo_margin!(style_mut);
    fo_min_height!(style_mut);
    fo_padding!(style_mut);
    style_dynamic_spacing!(style_mut);
    style_shadow!(style_mut);
    svg_height!(style_mut);
}
