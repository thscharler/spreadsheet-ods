use crate::attrmap2::AttrMap2;
use crate::style::units::{
    Angle, FontSize, FontStyle, FontVariant, FontWeight, Length, LetterSpacing, LineMode,
    LineStyle, LineType, LineWidth, Percent, RotationScale, TextCombine, TextCondition,
    TextDisplay, TextEmphasize, TextPosition, TextRelief, TextTransform,
};
use crate::style::Style;
use crate::style::{color_string, shadow_string, text_position, StyleOrigin, StyleUse};
use color::Rgb;
use icu_locid::Locale;
use std::fmt::{Display, Formatter};

style_ref!(TextStyleRef);

/// Text style.
/// This is not used for cell-formatting. Use CellStyle instead.
///
#[derive(Debug, Clone)]
pub struct TextStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// Style name
    name: String,
    /// General attributes
    // ok style:auto-update 19.467 => ALL
    // ok style:class 19.470, => ALL
    // ignore style:data-style-name 19.473, => CELL, CHART
    // ignore style:default-outlinelevel 19.474, => PARAGRAPH
    // ok style:display-name 19.476, => ALL
    // ignore style:family 19.480, => Not mapped as an attribute.
    // ignore style:list-level 19.499, => PARAGRAPH
    // ignore style:list-style-name 19.500, => PARAGRAPH
    // ignore style:master-page-name 19.501, => PARAGRAPH, TABLE
    // ignore style:name 19.502, => Not mapped as an attribute.
    // ignore style:next-style-name 19.503, => PARAGRAPH
    // ok style:parent-style-name 19.510 => ALL
    // ignore style:percentage-data-style-name 19.511. => PARAGRAPH?
    attr: AttrMap2,
    textstyle: AttrMap2,
}

styles_styles!(TextStyle, TextStyleRef);

impl TextStyle {
    /// Empty.
    pub(crate) fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            textstyle: Default::default(),
        }
    }

    /// A new named style.
    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            textstyle: Default::default(),
        }
    }

    /// General attributes for the style.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes for the style.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// All text-attributes for the style.
    pub(crate) fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// All text-attributes for the style.
    pub(crate) fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    fo_background_color!(textstyle_mut);
    fo_color!(textstyle_mut);
    fo_locale!(textstyle_mut);
    style_font_name!(textstyle_mut);
    fo_font_size!(textstyle_mut);
    fo_font_size_rel!(textstyle_mut);
    fo_font_style!(textstyle_mut);
    fo_font_weight!(textstyle_mut);
    fo_font_variant!(textstyle_mut);
    fo_font_attr!(textstyle_mut);
    style_locale_asian!(textstyle_mut);
    style_font_name_asian!(textstyle_mut);
    style_font_size_asian!(textstyle_mut);
    style_font_size_rel_asian!(textstyle_mut);
    style_font_style_asian!(textstyle_mut);
    style_font_weight_asian!(textstyle_mut);
    style_font_attr_asian!(textstyle_mut);
    style_locale_complex!(textstyle_mut);
    style_font_name_complex!(textstyle_mut);
    style_font_size_complex!(textstyle_mut);
    style_font_size_rel_complex!(textstyle_mut);
    style_font_style_complex!(textstyle_mut);
    style_font_weight_complex!(textstyle_mut);
    style_font_attr_complex!(textstyle_mut);
    fo_hyphenate!(textstyle_mut);
    fo_hyphenation_push_char_count!(textstyle_mut);
    fo_hyphenation_remain_char_count!(textstyle_mut);
    fo_letter_spacing!(textstyle_mut);
    fo_text_shadow!(textstyle_mut);
    fo_text_transform!(textstyle_mut);
    style_font_relief!(textstyle_mut);
    style_text_position!(textstyle_mut);
    style_rotation_angle!(textstyle_mut);
    style_rotation_scale!(textstyle_mut);
    style_letter_kerning!(textstyle_mut);
    style_text_combine!(textstyle_mut);
    style_text_combine_start_char!(textstyle_mut);
    style_text_combine_end_char!(textstyle_mut);
    style_text_emphasize!(textstyle_mut);
    style_text_line_through!(textstyle_mut);
    style_text_outline!(textstyle_mut);
    style_text_overline!(textstyle_mut);
    style_text_underline!(textstyle_mut);
    style_use_window_font_color!(textstyle_mut);
    text_condition!(textstyle_mut);
    text_display!(textstyle_mut);
}
