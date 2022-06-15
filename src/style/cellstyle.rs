use crate::attrmap2::AttrMap2;
use crate::format::ValueFormatRef;
use crate::style::stylemap::StyleMap;
use crate::style::units::{
    Angle, Border, CellAlignVertical, FontStyle, FontWeight, Length, LineMode, LineStyle, LineType,
    LineWidth, PageBreak, ParaAlignVertical, RotationAlign, TextAlign, TextAlignSource, TextKeep,
    TextPosition, TextRelief, TextTransform, WrapOption, WritingMode,
};
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string,
    StyleOrigin, StyleUse, TextStyleRef,
};
use color::Rgb;

style_ref!(CellStyleRef);

/// Describes the style information for a cell.
///
/// ```
/// use spreadsheet_ods::{pt, Length, CellStyle, WorkBook, Sheet, CellStyleRef};
/// use spreadsheet_ods::defaultstyles::DefaultFormat;
/// use color::Rgb;
///
/// let mut book = WorkBook::new();
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
    /// General attributes
    attr: AttrMap2,
    //
    cellstyle: AttrMap2,
    paragraphstyle: AttrMap2,
    textstyle: AttrMap2,
    /// Style maps
    stylemaps: Option<Vec<StyleMap>>,
}

impl CellStyle {
    /// Creates an empty style.
    pub fn empty() -> Self {
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

    /// Returns the name as a CellStyleRef.
    pub fn style_ref(&self) -> CellStyleRef {
        CellStyleRef::from(self.name())
    }

    /// Origin of the style, either styles.xml oder content.xml
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Changes the origin.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Usage for the style.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Usage for the style.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Stylename
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Stylename
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
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

    /// Display name.
    pub fn display_name(&self) -> Option<&String> {
        self.attr.attr("style:display-name")
    }

    /// Display name.
    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:display-name", name.into());
    }

    /// The parent style this derives from.
    pub fn parent_style(&self) -> Option<&String> {
        self.attr.attr("style:parent-style-name")
    }

    /// The parent style this derives from.
    pub fn set_parent_style(&mut self, name: &CellStyleRef) {
        self.attr
            .set_attr("style:parent-style-name", name.to_string());
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Allows access to all cell-style like attributes.
    pub fn cellstyle(&self) -> &AttrMap2 {
        &self.cellstyle
    }

    /// Allows access to all cell-style like attributes.
    pub fn cellstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.cellstyle
    }

    /// Allows access to all paragraph-style like attributes.
    pub fn paragraphstyle(&self) -> &AttrMap2 {
        &self.paragraphstyle
    }

    /// Allows access to all paragraph-style like attributes.
    pub fn paragraphstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.paragraphstyle
    }

    /// Allows access to all text-style like attributes.
    pub fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// Allows access to all text-style like attributes.
    pub fn textstyle_mut(&mut self) -> &mut AttrMap2 {
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

    fo_break!(paragraphstyle_mut);
    fo_keep_together!(paragraphstyle_mut);
    fo_keep_with_next!(paragraphstyle_mut);
    fo_margin!(paragraphstyle_mut);
    paragraph!(paragraphstyle_mut);

    text!(textstyle_mut);

    // missing:
    // style:cell-protect 20.253,
    // style:decimal-places 20.258,
    // style:direction 20.263,
    // style:glyph-orientation-vertical 20.297,
    // style:text-align-source 20.364,
    fo_background_color!(cellstyle_mut);
    fo_border!(cellstyle_mut);
    fo_padding!(cellstyle_mut);
    style_shadow!(cellstyle_mut);
    style_writing_mode!(cellstyle_mut);

    /// Wrap text.
    pub fn set_wrap_option(&mut self, wrap: WrapOption) {
        self.cellstyle.set_attr("fo:wrap-option", wrap.to_string());
    }

    /// Printing?
    pub fn set_print_content(&mut self, print: bool) {
        self.cellstyle
            .set_attr("style:print-content", print.to_string());
    }

    /// Repeat to fill.
    pub fn set_repeat_content(&mut self, print: bool) {
        self.cellstyle
            .set_attr("style:repeat-content", print.to_string());
    }

    /// Rotation
    pub fn set_rotation_align(&mut self, align: RotationAlign) {
        self.cellstyle
            .set_attr("style:rotation-align", align.to_string());
    }

    /// Rotation
    pub fn set_rotation_angle(&mut self, angle: Angle) {
        self.cellstyle
            .set_attr("style:rotation-angle", angle.to_string());
    }

    /// Shrink text to fit.
    pub fn set_shrink_to_fit(&mut self, shrink: bool) {
        self.cellstyle
            .set_attr("style:shrink-to-fit", shrink.to_string());
    }

    /// Vertical alignment.
    pub fn set_vertical_align(&mut self, align: CellAlignVertical) {
        self.cellstyle
            .set_attr("style:vertical-align", align.to_string());
    }

    /// Diagonal style.
    pub fn set_diagonal_bl_tr(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.cellstyle
            .set_attr("style:diagonal-bl-tr", border_string(width, border, color));
    }

    /// Widths for double borders.
    pub fn set_diagonal_bl_tr_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.cellstyle.set_attr(
            "style:diagonal-bl-tr-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Diagonal style.
    pub fn set_diagonal_tl_br(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.cellstyle
            .set_attr("style:diagonal-tl-br", border_string(width, border, color));
    }

    /// Widths for double borders.
    pub fn set_diagonal_tl_br_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.cellstyle.set_attr(
            "style:diagonal-tl-br-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }
}
