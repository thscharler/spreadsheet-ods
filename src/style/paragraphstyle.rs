use crate::attrmap2::AttrMap2;
use crate::io::parse::parse_u32;
use crate::style::tabstop::TabStop;
use crate::style::units::{
    Border, FontStyle, FontVariant, FontWeight, Hyphenation, HyphenationLadderCount, Indent,
    Length, LineBreak, LineHeight, LineMode, LineStyle, LineType, LineWidth, Margin, PageBreak,
    PageNumber, ParaAlignVertical, Percent, PunctuationWrap, RotationScale, TextAlign,
    TextAlignLast, TextAutoSpace, TextCombine, TextCondition, TextDisplay, TextEmphasize, TextKeep,
    TextPosition, TextRelief, TextTransform, WritingMode,
};
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string,
    text_position, Style, StyleOrigin, StyleUse, TextStyleRef,
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

    /// Within styles for paragraphs, style:next-style-name attribute specifies the style to be used
    /// for the next paragraph if a paragraph break is inserted in the user interface. By default, the current
    /// style is used as the next style.
    pub fn next_style(&self) -> Option<&String> {
        self.attr.attr("style:next-style-name")
    }

    /// Within styles for paragraphs, style:next-style-name attribute specifies the style to be used
    /// for the next paragraph if a paragraph break is inserted in the user interface. By default, the current
    /// style is used as the next style.
    pub fn set_next_style(&mut self, name: &ParagraphStyleRef) {
        self.attr
            .set_attr("style:next-style-name", name.to_string());
    }

    /// The style:default-outline-level attribute specifies a default outline level for a style with
    /// the style:family 19.480 attribute value paragraph.
    ///
    /// If the style:default-outline-level attribute is present in a paragraph style, and if this
    /// paragraph style is assigned to a paragraph or heading by user action, then the consumer should
    /// replace the paragraph or heading with a heading of the specified level, which has the same
    /// content and attributes as the original paragraph or heading.
    ///
    /// Note: This attribute does not modify the behavior of <text:p> 5.1.3 or
    /// <text:h> 5.1.2 elements, but only instructs a consumer to create one or the
    /// other when assigning a paragraph style as a result of user interface action while
    /// the document is edited.
    ///
    /// The style:default-outline-level attribute value can be empty. If empty, this attribute
    /// does not inherit a list style value from a parent style.
    pub fn default_outline_level(&self) -> Option<u32> {
        match self.attr.attr("style:default-outline-level") {
            None => None,
            Some(v) => match parse_u32(v.as_bytes()) {
                Ok(v) => Some(v),
                Err(_) => None,
            },
        }
    }

    /// The style:default-outline-level attribute specifies a default outline level for a style with
    /// the style:family 19.480 attribute value paragraph.
    pub fn set_default_outline_level(&mut self, level: u32) {
        self.attr
            .set_attr("style:default-outline-level", level.to_string());
    }

    /// The style:master-page-name attribute defines a master page for a paragraph or table style.
    /// This applies to automatic and common styles.
    ///
    /// If this attribute is associated with a style, a page break is inserted when the style is applied and
    /// the specified master page is applied to the resulting page.
    ///
    /// This attribute is ignored if it is associated with a paragraph style that is applied to a paragraph
    /// within a table.
    pub fn master_page_name(&self) -> Option<MasterPageRef> {
        self.attr
            .attr("style:master-page-name")
            .map(MasterPageRef::from)
    }

    /// The style:master-page-name attribute defines a master page for a paragraph or table style.
    /// This applies to automatic and common styles.
    pub fn set_master_page_name(&mut self, masterpage: &MasterPageRef) {
        self.attr
            .set_attr("style:master-page-name", masterpage.to_string());
    }

    /// Tabstops.
    pub fn add_tabstop(&mut self, ts: TabStop) {
        let tabstops = self.tabstops.get_or_insert_with(Vec::new);
        tabstops.push(ts);
    }

    /// Tabstops.
    pub fn tabstops(&self) -> Option<&Vec<TabStop>> {
        self.tabstops.as_ref()
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
    style_autotext_indent!(paragraphstyle_mut);
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
    // TODO: tab-stops

    text!(textstyle_mut);
    text_locale!(textstyle_mut);
    // style_rotation_angle!(textstyle_mut);
    style_rotation_scale!(textstyle_mut);
    // fo_background_color!(textstyle_mut);
}
