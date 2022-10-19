use crate::attrmap2::AttrMap2;
use crate::format::ValueFormatRef;
use crate::style::stylemap::StyleMap;
use crate::style::units::{
    Angle, Border, CellAlignVertical, CellProtect, FontStyle, FontVariant, FontWeight,
    GlyphOrientation, Hyphenation, HyphenationLadderCount, Indent, Length, LineBreak, LineHeight,
    LineMode, LineStyle, LineType, LineWidth, Margin, PageBreak, PageNumber, ParaAlignVertical,
    Percent, PunctuationWrap, RotationAlign, RotationScale, TextAlign, TextAlignLast,
    TextAlignSource, TextAutoSpace, TextCombine, TextCondition, TextDisplay, TextEmphasize,
    TextKeep, TextPosition, TextRelief, TextTransform, WrapOption, WritingDirection, WritingMode,
};
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string,
    text_position, Style, StyleOrigin, StyleUse, TextStyleRef,
};
use color::Rgb;
use icu_locid::Locale;
use std::fmt::{Display, Formatter};

style_ref!(CellStyleRef);

/// Describes the style information for a cell.
///
/// ```
/// use spreadsheet_ods::{pt, Length, CellStyle, WorkBook, Sheet, CellStyleRef};
/// use spreadsheet_ods::defaultstyles::DefaultFormat;
/// use color::Rgb;
/// use icu_locid::locale;
///
/// let mut book = WorkBook::new(locale!("en_US"));
///
/// let mut st_header = CellStyle::new("header", &DefaultFormat::default());
/// st_header.set_font_bold();
/// st_header.set_color(Rgb::new(255,255,0));
/// st_header.set_font_size(pt!(18));
/// let ref_header = book.add_cellstyle(st_header);
///
/// let mut sheet0 = Sheet::new("sheet 1");
/// sheet0.set_styled_value(0,0, "title", &ref_header);
///
/// // use a style defined later or elsewhere:
/// let ref_some = CellStyleRef::from("some_else");
/// sheet0.set_styled_value(1,0, "some", &ref_some);
///
/// ```
///
#[derive(Debug, Clone)]
pub struct CellStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// Style name.
    name: String,
    /// General attributes.
    attr: AttrMap2,
    /// Cell style attributes.
    cellstyle: AttrMap2,
    /// Paragraph style attributes.
    paragraphstyle: AttrMap2,
    /// Text style attributes.
    textstyle: AttrMap2,
    /// Style maps
    stylemaps: Option<Vec<StyleMap>>,
}

styles_styles!(CellStyle, CellStyleRef);

impl CellStyle {
    /// Creates an empty style.
    pub(crate) fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            cellstyle: Default::default(),
            paragraphstyle: Default::default(),
            textstyle: Default::default(),
            stylemaps: None,
        }
    }

    /// Creates an empty style with the given name and a reference to a
    /// value format.
    pub fn new<S: Into<String>>(name: S, value_format: &ValueFormatRef) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            cellstyle: Default::default(),
            paragraphstyle: Default::default(),
            textstyle: Default::default(),
            stylemaps: None,
        };
        s.set_value_format(value_format);
        s
    }

    /// Reference to the value format.
    pub fn value_format(&self) -> Option<&String> {
        self.attr.attr("style:data-style-name")
    }

    /// Reference to the value format.
    pub fn set_value_format(&mut self, name: &ValueFormatRef) {
        self.attr
            .set_attr("style:data-style-name", name.to_string());
    }

    /// Allows access to all attributes of the style itself.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Allows access to all attributes of the style itself.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Allows access to all cell-style like attributes.
    pub(crate) fn cellstyle(&self) -> &AttrMap2 {
        &self.cellstyle
    }

    /// Allows access to all cell-style like attributes.
    pub(crate) fn cellstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.cellstyle
    }

    /// Allows access to all paragraph-style like attributes.
    pub(crate) fn paragraphstyle(&self) -> &AttrMap2 {
        &self.paragraphstyle
    }

    /// Allows access to all paragraph-style like attributes.
    pub(crate) fn paragraphstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.paragraphstyle
    }

    /// Allows access to all text-style like attributes.
    pub(crate) fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// Allows access to all text-style like attributes.
    pub(crate) fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
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

    // Cell attributes.
    fo_background_color!(cellstyle_mut);
    fo_border!(cellstyle_mut);
    fo_padding!(cellstyle_mut);
    fo_wrap_option!(cellstyle_mut);
    fo_border_line_width!(cellstyle_mut);
    style_cell_protect!(cellstyle_mut);
    style_decimal_places!(cellstyle_mut);
    style_diagonal!(cellstyle_mut);
    style_direction!(cellstyle_mut);
    style_glyph_orientation_vertical!(cellstyle_mut);
    style_print_content!(cellstyle_mut);
    style_repeat_content!(cellstyle_mut);
    style_rotation_align!(cellstyle_mut);
    style_rotation_angle!(cellstyle_mut);
    style_shadow!(cellstyle_mut);
    style_shrink_to_fit!(cellstyle_mut);
    style_text_align_source!(cellstyle_mut);
    style_vertical_align!(cellstyle_mut);
    style_writing_mode!(cellstyle_mut);

    // Paragraph attributes.
    // TODO: Some attributes exist as both cell and as paragraph properties. Can't be mapped this way...

    // fo_background_color!(paragraphstyle_mut);
    // fo_border!(paragraphstyle_mut);
    fo_break!(paragraphstyle_mut);
    fo_hyphenation!(paragraphstyle_mut);
    fo_keep_together!(paragraphstyle_mut);
    fo_keep_with_next!(paragraphstyle_mut);
    fo_line_height!(paragraphstyle_mut);
    fo_margin!(paragraphstyle_mut);
    fo_orphans!(paragraphstyle_mut);
    // fo_padding!(paragraphstyle_mut);
    fo_text_align!(paragraphstyle_mut);
    fo_text_align_last!(paragraphstyle_mut);
    fo_text_indent!(paragraphstyle_mut);
    fo_widows!(paragraphstyle_mut);
    style_autotext_indent!(paragraphstyle_mut);
    style_background_transparency!(paragraphstyle_mut);
    // fo_border_line_width!(paragraphstyle_mut);
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
    // style_shadow!(paragraphstyle_mut);
    style_snap_to_layout_grid!(paragraphstyle_mut);
    style_tab_stop_distance!(paragraphstyle_mut);
    style_text_autospace!(paragraphstyle_mut);
    style_vertical_align_para!(paragraphstyle_mut);
    // style_writing_mode!(paragraphstyle_mut);
    style_writing_mode_automatic!(paragraphstyle_mut);
    style_line_number!(paragraphstyle_mut);
    style_number_lines!(paragraphstyle_mut);

    text!(textstyle_mut);
    text_locale!(textstyle_mut);
    // style_rotation_angle!(textstyle_mut);
    style_rotation_scale!(textstyle_mut);
    // fo_background_color!(textstyle_mut);

    // TODO: background image
}
