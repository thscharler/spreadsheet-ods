use crate::attrmap2::AttrMap2;
use crate::style::tabstop::TabStop;
use crate::style::units::{
    Border, FontSize, FontStyle, FontVariant, FontWeight, Hyphenation, HyphenationLadderCount,
    Indent, Length, LetterSpacing, LineBreak, LineHeight, LineMode, LineStyle, LineType, LineWidth,
    Margin, PageBreak, PageNumber, ParaAlignVertical, Percent, PunctuationWrap, RotationScale,
    TextAlign, TextAlignLast, TextAutoSpace, TextCombine, TextCondition, TextDisplay,
    TextEmphasize, TextKeep, TextPosition, TextRelief, TextTransform, WritingMode,
};
use crate::style::{
    border_line_width_string, border_string, color_string, shadow_string, text_position, Style,
    StyleOrigin, StyleUse, TextStyleRef,
};
use crate::MasterPageRef;
use color::Rgb;
use icu_locid::Locale;
use std::fmt::{Display, Formatter};

style_ref!(ParagraphStyleRef);

/// Paragraph style.
///
/// This is not used for cell-formatting. Use [crate::style::CellStyle] instead.
/// This kind of style is used for complex text formatting. See [crate::text].
///
#[derive(Debug, Clone)]
pub struct ParagraphStyle {
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
    // ok style:default-outlinelevel 19.474, => PARAGRAPH
    // ok style:display-name 19.476, => ALL
    // ignore style:family 19.480, => Not mapped as an attribute.
    // ignore style:list-level 19.499, => PARAGRAPH
    // ignore style:list-style-name 19.500, => PARAGRAPH
    // ok style:master-page-name 19.501, => PARAGRAPH, TABLE
    // ignore style:name 19.502, => Not mapped as an attribute.
    // ok style:next-style-name 19.503, => PARAGRAPH
    // ok style:parent-style-name 19.510 => ALL
    // ignore style:percentage-data-style-name 19.511. => PARAGRAPH?
    attr: AttrMap2,
    /// Paragraph attributes
    paragraphstyle: AttrMap2,
    /// Text attributes
    textstyle: AttrMap2,
    /// Tabstop data.
    tabstops: Option<Vec<TabStop>>,
}

styles_styles!(ParagraphStyle, ParagraphStyleRef);

impl ParagraphStyle {
    /// Empty
    pub(crate) fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            paragraphstyle: Default::default(),
            textstyle: Default::default(),
            tabstops: None,
        }
    }

    /// New style.
    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            paragraphstyle: Default::default(),
            textstyle: Default::default(),
            tabstops: None,
        }
    }

    /// General attributes.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Paragraph style attributes.
    pub(crate) fn paragraphstyle(&self) -> &AttrMap2 {
        &self.paragraphstyle
    }

    /// Paragraph style attributes.
    pub(crate) fn paragraphstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.paragraphstyle
    }

    /// Text style attributes.
    pub(crate) fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// Text style attributes.
    pub(crate) fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    style_default_outline_level!(attrmap_mut);
    style_master_page!(attrmap_mut);
    style_next_style!(attrmap_mut);

    /// Tabstops.
    pub fn add_tabstop(&mut self, ts: TabStop) {
        let tabstops = self.tabstops.get_or_insert_with(Vec::new);
        tabstops.push(ts);
    }

    /// Tabstops.
    pub fn tabstops(&self) -> Option<&Vec<TabStop>> {
        self.tabstops.as_ref()
    }

    fo_background_color!(paragraphstyle_mut);
    fo_border!(paragraphstyle_mut);
    fo_break!(paragraphstyle_mut);
    fo_hyphenation!(paragraphstyle_mut);
    fo_keep_together!(paragraphstyle_mut);
    fo_keep_with_next!(paragraphstyle_mut);
    fo_line_height!(paragraphstyle_mut);
    fo_margin!(paragraphstyle_mut);
    fo_orphans!(paragraphstyle_mut);
    fo_padding!(paragraphstyle_mut);
    fo_text_align!(paragraphstyle_mut);
    fo_text_align_last!(paragraphstyle_mut);
    fo_text_indent!(paragraphstyle_mut);
    fo_widows!(paragraphstyle_mut);
    style_auto_text_indent!(paragraphstyle_mut);
    style_background_transparency!(paragraphstyle_mut);
    fo_border_line_width!(paragraphstyle_mut);
    style_contextual_spacing!(paragraphstyle_mut);
    style_font_independent_line_spacing!(paragraphstyle_mut);
    style_join_border!(paragraphstyle_mut);
    style_justify_single_word!(paragraphstyle_mut);
    style_line_break!(paragraphstyle_mut);
    style_line_height_at_least!(paragraphstyle_mut);
    style_line_spacing!(paragraphstyle_mut);
    style_page_number!(paragraphstyle_mut);
    style_punctuation_wrap!(paragraphstyle_mut);
    style_register_true!(paragraphstyle_mut);
    style_shadow!(paragraphstyle_mut);
    style_snap_to_layout_grid!(paragraphstyle_mut);
    style_tab_stop_distance!(paragraphstyle_mut);
    style_text_autospace!(paragraphstyle_mut);
    style_vertical_align_para!(paragraphstyle_mut);
    style_writing_mode!(paragraphstyle_mut);
    style_writing_mode_automatic!(paragraphstyle_mut);
    style_line_number!(paragraphstyle_mut);
    style_number_lines!(paragraphstyle_mut);

    // TODO: background-image
    // TODO: drop-cap

    // fo_background_color!(textstyle_mut);
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
    // style_rotation_angle!(textstyle_mut);
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
