use crate::attrmap2::AttrMap2;
use crate::style::units::{
    Border, Margin, MasterPageUsage, Percent, PrintCentering, PrintContent, PrintOrder,
    PrintOrientation, StyleNumFormat, WritingMode,
};
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string,
    ParseStyleAttr,
};
use crate::{Length, OdsError};
use color::Rgb;
use std::fmt::{Display, Formatter};

style_ref!(PageStyleRef);

/// The <style:page-layout> element represents the styles that specify the formatting properties
/// of a page.
///
/// For an example see [MasterPage].
///
#[derive(Debug, Clone)]
pub struct PageStyle {
    name: String,
    // TODO: reading and writing work on strings, get/set on an enum. is this nice?
    pub(crate) master_page_usage: Option<String>,
    style: AttrMap2,
    header: HeaderFooterStyle,
    footer: HeaderFooterStyle,
}

impl PageStyle {
    /// New pagestyle.
    pub(crate) fn new_empty() -> Self {
        Self {
            name: Default::default(),
            master_page_usage: None,
            style: Default::default(),
            header: Default::default(),
            footer: Default::default(),
        }
    }

    /// New pagestyle.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            master_page_usage: None,
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

    /// The style:page-usage attribute specifies the type of pages that a page master should
    /// generate.
    /// The defined values for the style:page-usage attribute are:
    /// * all: if there are no <style:header-left> or <style:footer-left> elements, the
    /// header and footer content is the same for left and right pages.
    /// * left: <style:header-left> or <style:footer-left> elements are ignored.
    /// * mirrored: if there are no <style:header-left> or <style:footer-left> elements,
    /// the header and footer content is the same for left and right pages.
    /// * right: <style:header-left> or <style:footer-left> elements are ignored.
    ///
    /// The default value for this attribute is all.
    pub fn set_page_usage(&mut self, usage: Option<MasterPageUsage>) {
        self.master_page_usage = usage.map(|m| m.to_string());
    }

    /// The style:page-usage attribute specifies the type of pages that a page master should
    /// generate.
    pub fn page_usage(&self) -> Result<Option<MasterPageUsage>, OdsError> {
        MasterPageUsage::parse_attr(self.master_page_usage.as_ref())
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
    pub(crate) fn style(&self) -> &AttrMap2 {
        &self.style
    }

    /// Access to all style attributes.
    pub(crate) fn style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.style
    }

    // ok fo:background-color 20.175,
    // ok fo:border 20.176.2,
    // ok fo:border-bottom 20.176.3,
    // ok fo:border-left 20.176.4,
    // ok fo:border-right 20.176.5,
    // ok fo:border-top 20.176.6,
    // ok fo:margin 20.198,
    // ok fo:margin-bottom 20.199,
    // ok fo:margin-left 20.200,
    // ok fo:marginright 20.201,
    // ok fo:margin-top 20.202,
    // ok fo:padding 20.210,
    // ok fo:padding-bottom 20.211,
    // ok fo:padding-left 20.212,
    // ok fo:padding-right 20.213,
    // ok fo:padding-top 20.214,
    // ok fo:page-height 20.208,
    // ok fo:page-width 20.209,
    // ok style:border-line-width 20.241,
    // ok style:border-line-width-bottom 20.242,
    // ok style:border-line-width-left 20.243,
    // ok style:border-line-width-right 20.244,
    // ok style:border-line-width-top 20.245,
    // ok style:first-page-number 20.258,
    // ok style:footnote-max-height 20.288,
    // ignore style:layout-grid-base-height 20.296,
    // ignore style:layout-grid-base-width 20.297,
    // ignore style:layout-grid-color 20.298,
    // ignore style:layout-grid-display 20.299,
    // ignore style:layout-grid-lines 20.300,
    // ignore style:layout-grid-mode 20.301,
    // ignore style:layout-grid-print 20.302,
    // ignore style:layout-grid-ruby-below 20.303,
    // ignore style:layout-grid-ruby-height 20.304,
    // ignore style:layout-grid-snap-to 20.305,
    // ignore style:layout-grid-standard-mode 20.306,
    // ok style:num-format 20.314,
    // ok style:num-letter-sync 20.315,
    // ok style:num-prefix 20.316,
    // ok style:num-suffix 20.317,
    // ok style:paper-tray-name 20.321,
    // ok style:print 20.322,
    // ok style:print-orientation 20.325,
    // ok style:print-page-order 20.324,
    // ??? style:register-truth-ref-style-name 20.329,
    // ok style:scale-to 20.344,
    // ok style:scale-to-pages 20.345,
    // ok style:shadow 20.349,
    // ok style:table-centering 20.353
    // style:writing-mode 20.394.3

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

    /// The style:first-page-number attribute specifies the number of a document.
    /// The value of this attribute can be an integer or continue. If the value is continue, the page
    /// number is the preceding page number incremented by 1. The default first page number is 1.
    pub fn set_first_page_number(&mut self, number: u32) {
        self.style_mut()
            .set_attr("style:first-page-number", number.to_string());
    }

    /// The style:footnote-max-height attribute specifies the maximum amount of space on a
    /// page that a footnote can occupy. The value of the attribute is a length, which determines the
    /// maximum height of a footnote area.
    /// If the value of this attribute is set to 0cm, there is no limit to the amount of space that the footnote
    /// can occupy.
    pub fn set_footnote_max_height(&mut self, height: Length) {
        self.style_mut()
            .set_attr("style:footnote-max-height", height.to_string());
    }

    /// The style:num-format attribute specifies a numbering sequence.
    /// If no value is given, no number sequence is displayed.
    ///
    /// The defined values for the style:num-format attribute are:
    /// * 1: number sequence starts with “1”.
    /// * a: number sequence starts with “a”.
    /// * A: number sequence starts with “A”.
    /// * empty string: no number sequence displayed.
    /// * i: number sequence starts with “i”.
    /// * I: number sequence start with “I”.
    /// * a value of type string 18.2
    pub fn set_num_format(&mut self, format: StyleNumFormat) {
        self.style_mut()
            .set_attr("style:num-format", format.to_string());
    }

    /// The style:num-letter-sync attribute specifies whether letter synchronization shall take
    /// place. If letters are used in alphabetical order for numbering, there are two ways to process
    /// overflows within a digit, as follows:
    /// * false: A new digit is inserted that always has the same value as the following digit. The
    /// numbering sequence (for lower case numberings) in that case is a, b, c, ..., z, aa, bb, cc, ...,
    /// zz, aaa, ..., and so on.
    /// * true: A new digit is inserted. Its start value is ”a” or ”A”, and it is incremented every time an
    /// overflow occurs in the following digit. The numbering sequence (for lower case numberings) in
    /// that case is a,b,c, ..., z, aa, ab, ac, ...,az, ba, ..., and so on
    pub fn set_num_letter_sync(&mut self, sync: bool) {
        self.style_mut()
            .set_attr("style:num-letter-sync", sync.to_string());
    }

    /// The style:num-prefix attribute specifies what to display before a number.
    /// If the style:num-prefix and style:num-suffix values do not contain any character that
    /// has a Unicode category of Nd, Nl, No, Lu, Ll, Lt, Lm or Lo, an [XSLT] format attribute can be
    /// created from the OpenDocument attributes by concatenating the values of the style:num-prefix,
    /// style:num-format, and style:num-suffix attributes.
    pub fn set_num_prefix<S: Into<String>>(&mut self, prefix: S) {
        self.style_mut().set_attr("style:num-prefix", prefix.into());
    }

    /// The style:num-prefix and style:num-suffix attributes specify what to display before and
    /// after a number.
    /// If the style:num-prefix and style:num-suffix values do not contain any character that
    /// has a Unicode category of Nd, Nl, No, Lu, Ll, Lt, Lm or Lo, an [XSLT] format attribute can be
    /// created from the OpenDocument attributes by concatenating the values of the style:numprefix, style:num-format, and style:num-suffix attributes.
    pub fn set_num_suffix<S: Into<String>>(&mut self, suffix: S) {
        self.style_mut().set_attr("style:num-suffix", suffix.into());
    }

    /// The style:paper-tray-name attribute specifies the paper tray to use when printing a
    /// document. The names assigned to the paper trays depends upon the printer.
    /// The defined values for the style:paper-tray-name attribute are:
    /// * default: the default tray specified by printer configuration settings.
    /// * a value of type string
    pub fn set_paper_tray_name<S: Into<String>>(&mut self, tray: S) {
        self.style_mut()
            .set_attr("style:paper-tray-name", tray.into());
    }

    /// The style:print attribute specifies the components in a spreadsheet document to print.
    /// The value of the style:print attribute is a white space separated list of one or more of these
    /// values: headers, grid, annotations, objects, charts, drawings, formulas, zerovalues, or the empty list.
    /// The defined values for the style:print attribute are:
    /// * annotations: annotations should be printed.
    /// * charts: charts should be printed.
    /// * drawings: drawings should be printed.
    /// * formulas: formulas should be printed.
    /// * headers: headers should be printed.
    /// * grid: grid lines should be printed.
    /// * objects: (including graphics): objects should be printed.
    /// * zero-values: zero-values should be printed.
    pub fn set_print(&mut self, print: &[PrintContent]) {
        let mut buf = String::new();
        for p in print {
            buf.push_str(&p.to_string());
            buf.push(' ');
        }
        self.style_mut().set_attr("style:print", buf);
    }

    /// The style:print-orientation attribute specifies the orientation of the printed page. The
    /// value of this attribute can be portrait or landscape.
    /// The defined values for the style:print-orientation attribute are:
    /// * landscape: a page is printed in landscape orientation.
    /// * portrait: a page is printed in portrait orientation.
    pub fn set_print_orientation(&mut self, orientation: PrintOrientation) {
        self.style_mut()
            .set_attr("style:print-orientation", orientation.to_string());
    }

    /// The style:print-page-order attribute specifies the order in which data in a spreadsheet is
    /// numbered and printed when the data does not fit on one printed page.
    /// The defined values for the style:print-page-order attribute are:
    /// * ltr: create pages from the first column to the last column before continuing with the next set
    /// of rows.
    /// * ttb: create pages from the top row to the bottom row before continuing with the next set of
    /// columns.
    pub fn set_print_page_order(&mut self, order: PrintOrder) {
        self.style_mut()
            .set_attr("style:print-page-order", order.to_string());
    }

    /// The style:scale-to attribute specifies that a document is to be scaled to a percentage value.
    /// A value of 100% means no scaling.
    /// If this attribute and style:scale-to-pages are absent, a document is not scaled.
    pub fn set_scale_to(&mut self, percent: Percent) {
        self.style_mut()
            .set_attr("style:scale-to", percent.to_string());
    }

    /// The style:scale-to-pages attribute specifies the number of pages on which a document
    /// should be printed. The document is scaled to fit a specified number of pages.
    /// If this attribute and style:scale-to are absent, a document is not scaled.
    pub fn set_scale_to_pages(&mut self, pages: u32) {
        self.style_mut()
            .set_attr("style:scale-to-pages", pages.to_string());
    }

    /// The style:table-centering attribute specifies whether tables are centered horizontally
    /// and/or vertically on the page. This attribute only applies to spreadsheet documents.
    /// The default is to align the table to the top-left or top-right corner of the page, depending of its
    /// writing direction.
    /// The defined values for the style:table-centering attribute are:
    /// * both: tables should be centered both horizontally and vertically on the pages where they
    /// appear.
    /// * horizontal: tables should be centered horizontally on the pages where they appear.
    /// * none: tables should not be centered both horizontally or vertically on the pages where they
    /// appear.
    /// * vertical: tables should be centered vertically on the pages where they appear.
    pub fn set_table_centering(&mut self, center: PrintCentering) {
        self.style_mut()
            .set_attr("style:table-centering", center.to_string());
    }

    /// See §7.27.7 of [XSL].
    /// The defined value for the style:writing-mode attribute is page: writing mode is inherited from
    /// the page that contains the element where this attribute appears.
    pub fn set_writing_mode(&mut self, mode: WritingMode) {
        self.style_mut()
            .set_attr("style:writing-mode", mode.to_string());
    }

    fo_background_color!(style_mut);
    fo_border!(style_mut);
    fo_margin!(style_mut);
    fo_padding!(style_mut);
    style_dynamic_spacing!(style_mut);
    style_shadow!(style_mut);
}

/// Style attributes for header/footer.
#[derive(Clone, Debug, Default)]
pub struct HeaderFooterStyle {
    style: AttrMap2,
}

// ok fo:background-color 20.175,
// ok fo:border 20.176.2,
// ok fo:border-bottom 20.176.3,
// ok fo:border-left 20.176.4,
// ok fo:border-right 20.176.5,
// ok fo:border-top 20.176.6,
// ok fo:margin 20.198,
// ok fo:margin-bottom 20.199,
// ok fo:margin-left 20.200,
// ok fo:marginright 20.201,
// ok fo:margin-top 20.202,
// ok fo:min-height 20.205.2,
// ok fo:padding 20.210,
// ok fo:padding-bottom 20.211,
// ok fo:padding-left 20.212,
// ok fo:padding-right 20.213,
// ok fo:padding-top 20.214,
// ok style:border-line-width 20.241,
// ok style:border-line-width-bottom 20.242,
// ok style:border-line-width-left 20.243,
// ok style:border-line-width-right 20.244,
// ok style:border-line-width-top 20.245,
// ok style:dynamic-spacing 20.256,
// ok style:shadow 20.349
// ok svg:height 20.397.2.
impl HeaderFooterStyle {
    /// General attributes.
    pub(crate) fn style(&self) -> &AttrMap2 {
        &self.style
    }

    /// General attributes.
    pub(crate) fn style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.style
    }

    /// Height.
    pub fn set_height(&mut self, height: Length) {
        self.style_mut().set_attr("svg:height", height.to_string());
    }

    fo_background_color!(style_mut);
    fo_border!(style_mut);
    fo_margin!(style_mut);
    fo_min_height!(style_mut);
    fo_padding!(style_mut);
    fo_border_line_width!(style_mut);
    style_dynamic_spacing!(style_mut);
    style_shadow!(style_mut);
}
