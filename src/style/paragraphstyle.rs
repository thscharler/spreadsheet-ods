use crate::attrmap2::AttrMap2;
use crate::color::Rgb;
use crate::style::tabstop::TabStop;
use crate::style::units::{
    Border, FontSize, FontStyle, FontVariant, FontWeight, Hyphenation, HyphenationLadderCount,
    Indent, Length, LetterSpacing, LineBreak, LineHeight, LineMode, LineStyle, LineType, LineWidth,
    Margin, PageBreak, PageNumber, ParaAlignVertical, Percent, PunctuationWrap, RotationScale,
    TextAlign, TextAlignLast, TextAutoSpace, TextCombine, TextCondition, TextDisplay,
    TextEmphasize, TextEmphasizePosition, TextKeep, TextPosition, TextRelief, TextTransform,
    WritingMode,
};
use crate::style::AnyStyleRef;
use crate::style::MasterPageRef;
use crate::style::{
    border_line_width_string, border_string, color_string, shadow_string, text_position,
    StyleOrigin, StyleUse, TextStyleRef,
};
use get_size2::GetSize;
use icu_locid::Locale;
use std::borrow::Borrow;

style_ref2!(ParagraphStyleRef);

/// Paragraph style.
///
/// This is not used for cell-formatting. Use [crate::style::CellStyle] instead.
/// This kind of style is used for complex text formatting. See [crate::text].
///
#[derive(Debug, Clone, GetSize)]
pub struct ParagraphStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// Style name
    name: String,
    /// General attributes
    attr: AttrMap2,
    /// Paragraph attributes
    paragraphstyle: AttrMap2,
    /// Text attributes
    textstyle: AttrMap2,
    /// Tabstop data.
    tabstops: Option<Vec<TabStop>>,
}

styles_styles2!(ParagraphStyle, ParagraphStyleRef);

impl ParagraphStyle {
    /// Empty
    pub fn new_empty() -> Self {
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
    pub fn new<S: AsRef<str>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.as_ref().to_string(),
            attr: Default::default(),
            paragraphstyle: Default::default(),
            textstyle: Default::default(),
            tabstops: None,
        }
    }

    /// General attributes.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Paragraph style attributes.
    pub fn paragraphstyle(&self) -> &AttrMap2 {
        &self.paragraphstyle
    }

    /// Paragraph style attributes.
    pub fn paragraphstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.paragraphstyle
    }

    /// Text style attributes.
    pub fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// Text style attributes.
    pub fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    style_default_outline_level!(attr);
    style_master_page!(attr);
    style_next_style!(attr);

    /// Tabstops.
    pub fn add_tabstop(&mut self, ts: TabStop) {
        let tabstops = self.tabstops.get_or_insert_with(Vec::new);
        tabstops.push(ts);
    }

    /// Tabstops.
    pub fn tabstops(&self) -> Option<&Vec<TabStop>> {
        self.tabstops.as_ref()
    }

    fo_background_color!(paragraphstyle);
    fo_border!(paragraphstyle);
    fo_break!(paragraphstyle);
    fo_hyphenation!(paragraphstyle);
    fo_keep_together!(paragraphstyle);
    fo_keep_with_next!(paragraphstyle);
    fo_line_height!(paragraphstyle);
    fo_margin!(paragraphstyle);
    fo_orphans!(paragraphstyle);
    fo_padding!(paragraphstyle);
    fo_text_align!(paragraphstyle);
    fo_text_align_last!(paragraphstyle);
    fo_text_indent!(paragraphstyle);
    fo_widows!(paragraphstyle);
    style_auto_text_indent!(paragraphstyle);
    style_background_transparency!(paragraphstyle);
    fo_border_line_width!(paragraphstyle);
    style_contextual_spacing!(paragraphstyle);
    style_font_independent_line_spacing!(paragraphstyle);
    style_join_border!(paragraphstyle);
    style_justify_single_word!(paragraphstyle);
    style_line_break!(paragraphstyle);
    style_line_height_at_least!(paragraphstyle);
    style_line_spacing!(paragraphstyle);
    style_page_number!(paragraphstyle);
    style_punctuation_wrap!(paragraphstyle);
    style_register_true!(paragraphstyle);
    style_shadow!(paragraphstyle);
    style_snap_to_layout_grid!(paragraphstyle);
    style_tab_stop_distance!(paragraphstyle);
    style_text_autospace!(paragraphstyle);
    style_vertical_align_para!(paragraphstyle);
    style_writing_mode!(paragraphstyle);
    style_writing_mode_automatic!(paragraphstyle);
    style_line_number!(paragraphstyle);
    style_number_lines!(paragraphstyle);

    // TODO: background-image
    // TODO: drop-cap

    // fo_background_color!(textstyle);
    fo_color!(textstyle);
    fo_locale!(textstyle);
    style_font_name!(textstyle);
    fo_font_size!(textstyle);
    fo_font_size_rel!(textstyle);
    fo_font_style!(textstyle);
    fo_font_weight!(textstyle);
    fo_font_variant!(textstyle);
    fo_font_attr!(textstyle);
    style_locale_asian!(textstyle);
    style_font_name_asian!(textstyle);
    style_font_size_asian!(textstyle);
    style_font_size_rel_asian!(textstyle);
    style_font_style_asian!(textstyle);
    style_font_weight_asian!(textstyle);
    style_font_attr_asian!(textstyle);
    style_locale_complex!(textstyle);
    style_font_name_complex!(textstyle);
    style_font_size_complex!(textstyle);
    style_font_size_rel_complex!(textstyle);
    style_font_style_complex!(textstyle);
    style_font_weight_complex!(textstyle);
    style_font_attr_complex!(textstyle);
    fo_hyphenate!(textstyle);
    fo_hyphenation_push_char_count!(textstyle);
    fo_hyphenation_remain_char_count!(textstyle);
    fo_letter_spacing!(textstyle);
    fo_text_shadow!(textstyle);
    fo_text_transform!(textstyle);
    style_font_relief!(textstyle);
    style_text_position!(textstyle);
    // style_rotation_angle!(textstyle);
    style_rotation_scale!(textstyle);
    style_letter_kerning!(textstyle);
    style_text_combine!(textstyle);
    style_text_combine_start_char!(textstyle);
    style_text_combine_end_char!(textstyle);
    style_text_emphasize!(textstyle);
    style_text_line_through!(textstyle);
    style_text_outline!(textstyle);
    style_text_overline!(textstyle);
    style_text_underline!(textstyle);
    style_use_window_font_color!(textstyle);
    text_condition!(textstyle);
    text_display!(textstyle);
}
