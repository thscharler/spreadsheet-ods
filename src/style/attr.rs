use crate::attrmap::AttrMap;
use crate::style::{
    border_line_width_string, border_string, color_string, percent_string, shadow_string, Border,
    CellAlignVertical, FontPitch, FontStyle, FontWeight, LineMode, LineStyle, LineType, LineWidth,
    PageBreak, ParaAlignVertical, RotationAlign, TextAlign, TextAlignSource, TextKeep,
    TextPosition, TextRelief, TextTransform, WrapOption, WritingMode,
};
use crate::{style, Angle, Length};
use color::Rgb;

/// Attributes for a FontFaceDecl
pub trait AttrFontDecl
where
    Self: AttrMap,
{
    /// Font-name
    fn set_name<S: Into<String>>(&mut self, name: S) {
        self.set_attr("style:name", name.into());
    }

    /// External font family name.
    fn set_font_family<S: Into<String>>(&mut self, name: S) {
        self.set_attr("svg:font-family", name.into());
    }

    /// System generic name.
    fn set_font_family_generic<S: Into<String>>(&mut self, name: S) {
        self.set_attr("style:font-family-generic", name.into());
    }

    /// Font pitch.
    fn set_font_pitch(&mut self, pitch: FontPitch) {
        self.set_attr("style:font-pitch", pitch.to_string());
    }
}

/// Margin attributes.
pub trait AttrFoMargin
where
    Self: AttrMap,
{
    fn set_margin(&mut self, margin: Length) {
        self.set_attr("fo:margin", margin.to_string());
    }

    fn set_margin_bottom(&mut self, margin: Length) {
        self.set_attr("fo:margin-bottom", margin.to_string());
    }

    fn set_margin_left(&mut self, margin: Length) {
        self.set_attr("fo:margin-left", margin.to_string());
    }

    fn set_margin_right(&mut self, margin: Length) {
        self.set_attr("fo:margin-right", margin.to_string());
    }

    fn set_margin_top(&mut self, margin: Length) {
        self.set_attr("fo:margin-top", margin.to_string());
    }
}

/// Padding attributes.
pub trait AttrFoPadding
where
    Self: AttrMap,
{
    fn set_padding(&mut self, padding: Length) {
        self.set_attr("fo:padding", padding.to_string());
    }

    fn set_padding_bottom(&mut self, padding: Length) {
        self.set_attr("fo:padding-bottom", padding.to_string());
    }

    fn set_padding_left(&mut self, padding: Length) {
        self.set_attr("fo:padding-left", padding.to_string());
    }

    fn set_padding_right(&mut self, padding: Length) {
        self.set_attr("fo:padding-right", padding.to_string());
    }

    fn set_padding_top(&mut self, padding: Length) {
        self.set_attr("fo:padding-top", padding.to_string());
    }
}

/// Background color.
pub trait AttrFoBackgroundColor
where
    Self: AttrMap,
{
    /// Border style.
    fn set_background_color(&mut self, color: Rgb<u8>) {
        self.set_attr("fo:background-color", color_string(color));
    }
}

/// Minimum height.
pub trait AttrFoMinHeight
where
    Self: AttrMap,
{
    fn set_min_height(&mut self, height: Length) {
        self.set_attr("fo:min-height", height.to_string());
    }

    fn set_min_height_percent(&mut self, height: f64) {
        self.set_attr("fo:min-height", percent_string(height));
    }
}

/// Border attributes.
pub trait AttrFoBorder
where
    Self: AttrMap,
{
    /// Border style all four sides.
    fn set_border(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_bottom(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-bottom", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_top(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-top", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_left(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-left", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_right(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-right", border_string(width, border, color));
    }

    /// Widths for double borders.
    fn set_border_line_width(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_bottom(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-bottom",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_left(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-left",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_right(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-right",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_top(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-top",
            border_line_width_string(inner, spacing, outer),
        );
    }
}

/// Page breaks.
pub trait AttrFoBreak
where
    Self: AttrMap,
{
    /// page-break
    fn set_break_before(&mut self, pagebreak: PageBreak) {
        self.set_attr("fo:break-before", pagebreak.to_string());
    }

    // page-break
    fn set_break_after(&mut self, pagebreak: PageBreak) {
        self.set_attr("fo:break-after", pagebreak.to_string());
    }
}

/// Keep with next.
pub trait AttrFoKeepWithNext
where
    Self: AttrMap,
{
    /// page-break
    fn set_keep_with_next(&mut self, keep_with_next: TextKeep) {
        self.set_attr("fo:keep-with-next", keep_with_next.to_string());
    }
}

/// Keep together.
pub trait AttrFoKeepTogether
where
    Self: AttrMap,
{
    /// page-break
    fn set_keep_together(&mut self, keep_together: TextKeep) {
        self.set_attr("fo:keep-together", keep_together.to_string());
    }
}

/// Height attribute.
pub trait AttrSvgHeight
where
    Self: AttrMap,
{
    fn set_height(&mut self, height: Length) {
        self.set_attr("svg:height", height.to_string());
    }
}

/// Spacing for header/footer.
pub trait AttrStyleDynamicSpacing
where
    Self: AttrMap,
{
    fn set_dynamic_spacing(&mut self, dynamic: bool) {
        self.set_attr("style:dynamic-spacing", dynamic.to_string());
    }
}

/// Shadows. Only a single shadow supported here.
pub trait AttrStyleShadow
where
    Self: AttrMap,
{
    fn set_shadow(
        &mut self,
        x_offset: Length,
        y_offset: Length,
        blur: Option<Length>,
        color: Rgb<u8>,
    ) {
        self.set_attr(
            "style:shadow",
            shadow_string(x_offset, y_offset, blur, color),
        );
    }
}

/// Writing mode.
pub trait AttrStyleWritingMode
where
    Self: AttrMap,
{
    fn set_writing_mode(&mut self, writing_mode: WritingMode) {
        self.set_attr("style:writing-mode", writing_mode.to_string());
    }
}

/// Table row specific attributes.
pub trait AttrTableRow
where
    Self: AttrMap,
{
    fn set_min_row_height(&mut self, min_height: Length) {
        self.set_attr("style:min-row-height", min_height.to_string());
    }

    fn set_row_height(&mut self, height: Length) {
        self.set_attr("style:row-height", height.to_string());
    }

    fn set_use_optimal_row_height(&mut self, opt: bool) {
        self.set_attr("style:use-optimal-row-height", opt.to_string());
    }
}

/// Table columns specific attributes.
pub trait AttrTableCol
where
    Self: AttrMap,
{
    /// Relative weights for the column width
    fn set_rel_col_width(&mut self, rel: f64) {
        self.set_attr("style:rel-column-width", style::rel_width_string(rel));
    }

    /// Column width
    fn set_col_width(&mut self, width: Length) {
        self.set_attr("style:column-width", width.to_string());
    }

    /// Override switch for the column width.
    fn set_use_optimal_col_width(&mut self, opt: bool) {
        self.set_attr("style:use-optimal-column-width", opt.to_string());
    }
}

/// Table cell specific styles.
pub trait AttrTableCell
where
    Self: AttrMap,
{
    fn set_wrap_option(&mut self, wrap: WrapOption) {
        self.set_attr("fo:wrap-option", wrap.to_string());
    }

    fn set_print_content(&mut self, print: bool) {
        self.set_attr("style:print-content", print.to_string());
    }

    fn set_repeat_content(&mut self, print: bool) {
        self.set_attr("style:repeat-content", print.to_string());
    }

    fn set_rotation_align(&mut self, align: RotationAlign) {
        self.set_attr("style:rotation-align", align.to_string());
    }

    fn set_rotation_angle(&mut self, angle: Angle) {
        self.set_attr("style:rotation-angle", angle.to_string());
    }

    fn set_shrink_to_fit(&mut self, shrink: bool) {
        self.set_attr("style:shrink-to-fit", shrink.to_string());
    }

    fn set_vertical_align(&mut self, align: CellAlignVertical) {
        self.set_attr("style:vertical-align", align.to_string());
    }

    /// Diagonal style.
    fn set_diagonal_bl_tr(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("style:diagonal-bl-tr", border_string(width, border, color));
    }

    /// Widths for double borders.
    fn set_diagonal_bl_tr_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:diagonal-bl-tr-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Diagonal style.
    fn set_diagonal_tl_br(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("style:diagonal-tl-br", border_string(width, border, color));
    }

    /// Widths for double borders.
    fn set_diagonal_tl_br_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:diagonal-tl-br-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }
}

/// Paragraph specific styles.
pub trait AttrParagraph
where
    Self: AttrMap,
{
    fn set_text_align_source(&mut self, align: TextAlignSource) {
        self.set_attr("style:text-align-source", align.to_string());
    }

    fn set_text_align(&mut self, align: TextAlign) {
        self.set_attr("fo:text-align", align.to_string());
    }

    fn set_text_indent(&mut self, indent: Length) {
        self.set_attr("fo:text-indent", indent.to_string());
    }

    fn set_line_spacing(&mut self, spacing: Length) {
        self.set_attr("style:line-spacing", spacing.to_string());
    }

    fn set_number_lines(&mut self, number: bool) {
        self.set_attr("text:number-lines", number.to_string());
    }

    fn set_vertical_align(&mut self, align: ParaAlignVertical) {
        self.set_attr("style:vertical-align", align.to_string());
    }
}

/// Text style attributes.
pub trait AttrText
where
    Self: AttrMap,
{
    fn set_color(&mut self, color: Rgb<u8>) {
        self.set_attr("fo:color", color_string(color));
    }

    fn set_font_name<S: Into<String>>(&mut self, name: S) {
        self.set_attr("style:font-name", name.into());
    }

    fn set_font_attr(&mut self, size: Length, bold: bool, italic: bool) {
        self.set_font_size(size);
        if bold {
            self.set_font_italic();
        }
        if italic {
            self.set_font_bold();
        }
    }

    fn set_font_size(&mut self, size: Length) {
        self.set_attr("fo:font-size", size.to_string());
    }

    fn set_font_size_percent(&mut self, size: f64) {
        self.set_attr("fo:font-size", percent_string(size));
    }

    fn set_font_italic(&mut self) {
        self.set_attr("fo:font-style", "italic".to_string());
    }

    fn set_font_style(&mut self, style: FontStyle) {
        self.set_attr("fo:font-style", style.to_string());
    }

    fn set_font_bold(&mut self) {
        self.set_attr("fo:font-weight", FontWeight::Bold.to_string());
    }

    fn set_font_weight(&mut self, weight: FontWeight) {
        self.set_attr("fo:font-weight", weight.to_string());
    }

    fn set_letter_spacing(&mut self, spacing: Length) {
        self.set_attr("fo:letter-spacing", spacing.to_string());
    }

    fn set_letter_spacing_normal(&mut self) {
        self.set_attr("fo:letter-spacing", "normal".to_string());
    }

    fn set_text_shadow(
        &mut self,
        x_offset: Length,
        y_offset: Length,
        blur: Option<Length>,
        color: Rgb<u8>,
    ) {
        self.set_attr(
            "fo:text-shadow",
            shadow_string(x_offset, y_offset, blur, color),
        );
    }

    fn set_text_position(&mut self, pos: TextPosition) {
        self.set_attr("style:text-position", pos.to_string());
    }

    fn set_text_transform(&mut self, trans: TextTransform) {
        self.set_attr("fo:text-transform", trans.to_string());
    }

    fn set_font_relief(&mut self, relief: TextRelief) {
        self.set_attr("style:font-relief", relief.to_string());
    }

    fn set_font_line_through_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:text-line-through-color", color_string(color));
    }

    fn set_font_line_through_style(&mut self, lstyle: LineStyle) {
        self.set_attr("style:text-line-through-style", lstyle.to_string());
    }

    fn set_font_line_through_mode(&mut self, lmode: LineMode) {
        self.set_attr("style:text-line-through-mode", lmode.to_string());
    }

    fn set_font_line_through_type(&mut self, ltype: LineType) {
        self.set_attr("style:text-line-through-type", ltype.to_string());
    }

    fn set_font_line_through_text<S: Into<String>>(&mut self, text: S) {
        self.set_attr("style:text-line-through-text", text.into());
    }

    fn set_font_line_through_text_style<S: Into<String>>(&mut self, style_ref: S) {
        self.set_attr("style:text-line-through-text-style", style_ref.into());
    }

    fn set_font_line_through_width(&mut self, lwidth: LineWidth) {
        self.set_attr("style:text-line-through-width", lwidth.to_string());
    }

    fn set_font_text_outline(&mut self, outline: bool) {
        self.set_attr("style:text-outline", outline.to_string());
    }

    fn set_font_underline_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:text-underline-color", color_string(color));
    }

    fn set_font_underline_style(&mut self, lstyle: LineStyle) {
        self.set_attr("style:text-underline-style", lstyle.to_string());
    }

    fn set_font_underline_type(&mut self, ltype: LineType) {
        self.set_attr("style:text-underline-type", ltype.to_string());
    }

    fn set_font_underline_mode(&mut self, lmode: LineMode) {
        self.set_attr("style:text-underline-mode", lmode.to_string());
    }

    fn set_font_underline_width(&mut self, lwidth: LineWidth) {
        self.set_attr("style:text-underline-width", lwidth.to_string());
    }

    fn set_font_overline_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:text-overline-color", color_string(color));
    }

    fn set_font_overline_style(&mut self, lstyle: LineStyle) {
        self.set_attr("style:text-overline-style", lstyle.to_string());
    }

    fn set_font_overline_type(&mut self, ltype: LineType) {
        self.set_attr("style:text-overline-type", ltype.to_string());
    }

    fn set_font_overline_mode(&mut self, lmode: LineMode) {
        self.set_attr("style:text-overline-mode", lmode.to_string());
    }

    fn set_font_overline_width(&mut self, lwidth: LineWidth) {
        self.set_attr("style:text-overline-width", lwidth.to_string());
    }
}
