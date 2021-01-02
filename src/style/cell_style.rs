use crate::attrmap2::AttrMap2;
use crate::style::stylemap::StyleMap;
use crate::style::units::{
    Angle, Border, CellAlignVertical, FontStyle, FontWeight, Length, LineMode, LineStyle, LineType,
    LineWidth, PageBreak, ParaAlignVertical, RotationAlign, TextAlign, TextAlignSource, TextKeep,
    TextPosition, TextRelief, TextTransform, WrapOption, WritingMode,
};
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string,
    StyleOrigin, StyleUse,
};
use color::Rgb;

#[derive(Debug, Clone)]
pub struct CellStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// General attributes
    // ??? style:auto-update 19.467,
    // ??? style:class 19.470,
    // ok style:data-style-name 19.473,
    // ignore style:default-outlinelevel 19.474,
    // ok style:display-name 19.476,
    // ok style:family 19.480,
    // ignore style:list-level 19.499,
    // ignore style:list-style-name 19.500,
    // ignore style:master-page-name 19.501,
    // ok style:name 19.502,
    // ignore style:next-style-name 19.503,
    // ok style:parent-style-name 19.510,
    // ignore style:percentage-data-style-name 19.511.
    attr: AttrMap2,
    cell_style: AttrMap2,
    paragraph_style: AttrMap2,
    text_style: AttrMap2,
    /// Style maps
    stylemaps: Option<Vec<StyleMap>>,
}

impl CellStyle {
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            cell_style: Default::default(),
            paragraph_style: Default::default(),
            text_style: Default::default(),
            stylemaps: None,
        }
    }

    pub fn new<S: Into<String>, T: Into<String>>(name: S, value_format: T) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            cell_style: Default::default(),
            paragraph_style: Default::default(),
            text_style: Default::default(),
            stylemaps: None,
        };
        s.set_name(name.into());
        s.set_value_format(value_format.into());
        s
    }

    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    pub fn name(&self) -> Option<&String> {
        self.attr.attr("style:name")
    }

    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:name", name.into());
    }

    pub fn value_format(&self) -> Option<&String> {
        self.attr.attr("style:data-style-name")
    }

    pub fn set_value_format<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:data-style-name", name.into());
    }

    pub fn display_name(&self) -> Option<&String> {
        self.attr.attr("style:display-name")
    }

    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:display-name", name.into());
    }

    pub fn parent_style(&self) -> Option<&String> {
        self.attr.attr("style:parent-style-name")
    }

    pub fn set_parent_style<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:parent-style-name", name.into());
    }

    pub fn attr_map(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attr_map_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    pub fn cell_style(&self) -> &AttrMap2 {
        &self.cell_style
    }

    pub fn cell_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.cell_style
    }

    pub fn paragraph_style(&self) -> &AttrMap2 {
        &self.paragraph_style
    }

    pub fn paragraph_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.paragraph_style
    }

    pub fn text_style(&self) -> &AttrMap2 {
        &self.text_style
    }

    pub fn text_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.text_style
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

    fo_background_color!(cell_style_mut);
    fo_border!(cell_style_mut);
    fo_padding!(cell_style_mut);
    style_shadow!(cell_style_mut);
    style_writing_mode!(cell_style_mut);

    fo_break!(paragraph_style_mut);
    fo_keep_together!(paragraph_style_mut);
    fo_keep_with_next!(paragraph_style_mut);
    fo_margin!(paragraph_style_mut);
    paragraph!(paragraph_style_mut);

    text!(text_style_mut);

    pub fn set_wrap_option(&mut self, wrap: WrapOption) {
        self.cell_style.set_attr("fo:wrap-option", wrap.to_string());
    }

    pub fn set_print_content(&mut self, print: bool) {
        self.cell_style
            .set_attr("style:print-content", print.to_string());
    }

    pub fn set_repeat_content(&mut self, print: bool) {
        self.cell_style
            .set_attr("style:repeat-content", print.to_string());
    }

    pub fn set_rotation_align(&mut self, align: RotationAlign) {
        self.cell_style
            .set_attr("style:rotation-align", align.to_string());
    }

    pub fn set_rotation_angle(&mut self, angle: Angle) {
        self.cell_style
            .set_attr("style:rotation-angle", angle.to_string());
    }

    pub fn set_shrink_to_fit(&mut self, shrink: bool) {
        self.cell_style
            .set_attr("style:shrink-to-fit", shrink.to_string());
    }

    pub fn set_vertical_align(&mut self, align: CellAlignVertical) {
        self.cell_style
            .set_attr("style:vertical-align", align.to_string());
    }

    /// Diagonal style.
    pub fn set_diagonal_bl_tr(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.cell_style
            .set_attr("style:diagonal-bl-tr", border_string(width, border, color));
    }

    /// Widths for double borders.
    pub fn set_diagonal_bl_tr_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.cell_style.set_attr(
            "style:diagonal-bl-tr-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Diagonal style.
    pub fn set_diagonal_tl_br(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.cell_style
            .set_attr("style:diagonal-tl-br", border_string(width, border, color));
    }

    /// Widths for double borders.
    pub fn set_diagonal_tl_br_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.cell_style.set_attr(
            "style:diagonal-tl-br-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }
}
