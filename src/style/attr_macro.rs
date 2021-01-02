#[macro_export]
macro_rules! fo_background_color {
    ($acc:ident) => {
        /// Border style.
        pub fn set_background_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:background-color", color_string(color));
        }
    };
}

#[macro_export]
macro_rules! fo_break {
    ($acc:ident) => {
        /// page-break
        pub fn set_break_before(&mut self, pagebreak: PageBreak) {
            self.$acc()
                .set_attr("fo:break-before", pagebreak.to_string());
        }

        // page-break
        pub fn set_break_after(&mut self, pagebreak: PageBreak) {
            self.$acc()
                .set_attr("fo:break-after", pagebreak.to_string());
        }
    };
}

#[macro_export]
macro_rules! fo_keep_with_next {
    ($acc:ident) => {
        /// page-break
        pub fn set_keep_with_next(&mut self, keep_with_next: TextKeep) {
            self.$acc()
                .set_attr("fo:keep-with-next", keep_with_next.to_string());
        }
    };
}

#[macro_export]
macro_rules! style_shadow {
    ($acc:ident) => {
        pub fn set_shadow(
            &mut self,
            x_offset: Length,
            y_offset: Length,
            blur: Option<Length>,
            color: Rgb<u8>,
        ) {
            self.$acc().set_attr(
                "style:shadow",
                shadow_string(x_offset, y_offset, blur, color),
            );
        }
    };
}

#[macro_export]
macro_rules! style_writing_mode {
    ($acc:ident) => {
        pub fn set_writing_mode(&mut self, writing_mode: WritingMode) {
            self.$acc()
                .set_attr("style:writing-mode", writing_mode.to_string());
        }
    };
}

#[macro_export]
macro_rules! fo_keep_together {
    ($acc:ident) => {
        /// page-break
        pub fn set_keep_together(&mut self, keep_together: TextKeep) {
            self.$acc()
                .set_attr("fo:keep-together", keep_together.to_string());
        }
    };
}

#[macro_export]
macro_rules! fo_border {
    ($acc:ident) => {
        /// Border style all four sides.
        pub fn set_border(&mut self, width: Length, border: Border, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:border", border_string(width, border, color));
        }

        /// Border style.
        pub fn set_border_bottom(&mut self, width: Length, border: Border, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:border-bottom", border_string(width, border, color));
        }

        /// Border style.
        pub fn set_border_top(&mut self, width: Length, border: Border, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:border-top", border_string(width, border, color));
        }

        /// Border style.
        pub fn set_border_left(&mut self, width: Length, border: Border, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:border-left", border_string(width, border, color));
        }

        /// Border style.
        pub fn set_border_right(&mut self, width: Length, border: Border, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:border-right", border_string(width, border, color));
        }

        /// Widths for double borders.
        pub fn set_border_line_width(&mut self, inner: Length, spacing: Length, outer: Length) {
            self.$acc().set_attr(
                "style:border-line-width",
                border_line_width_string(inner, spacing, outer),
            );
        }

        /// Widths for double borders.
        pub fn set_border_line_width_bottom(
            &mut self,
            inner: Length,
            spacing: Length,
            outer: Length,
        ) {
            self.$acc().set_attr(
                "style:border-line-width-bottom",
                border_line_width_string(inner, spacing, outer),
            );
        }

        /// Widths for double borders.
        pub fn set_border_line_width_left(
            &mut self,
            inner: Length,
            spacing: Length,
            outer: Length,
        ) {
            self.$acc().set_attr(
                "style:border-line-width-left",
                border_line_width_string(inner, spacing, outer),
            );
        }

        /// Widths for double borders.
        pub fn set_border_line_width_right(
            &mut self,
            inner: Length,
            spacing: Length,
            outer: Length,
        ) {
            self.$acc().set_attr(
                "style:border-line-width-right",
                border_line_width_string(inner, spacing, outer),
            );
        }

        /// Widths for double borders.
        pub fn set_border_line_width_top(&mut self, inner: Length, spacing: Length, outer: Length) {
            self.$acc().set_attr(
                "style:border-line-width-top",
                border_line_width_string(inner, spacing, outer),
            );
        }
    };
}

#[macro_export]
macro_rules! fo_padding {
    ($acc:ident) => {
        pub fn set_padding(&mut self, padding: Length) {
            self.$acc().set_attr("fo:padding", padding.to_string());
        }

        pub fn set_padding_bottom(&mut self, padding: Length) {
            self.$acc()
                .set_attr("fo:padding-bottom", padding.to_string());
        }

        pub fn set_padding_left(&mut self, padding: Length) {
            self.$acc().set_attr("fo:padding-left", padding.to_string());
        }

        pub fn set_padding_right(&mut self, padding: Length) {
            self.$acc()
                .set_attr("fo:padding-right", padding.to_string());
        }

        pub fn set_padding_top(&mut self, padding: Length) {
            self.$acc().set_attr("fo:padding-top", padding.to_string());
        }
    };
}

#[macro_export]
macro_rules! fo_margin {
    ($acc:ident) => {
        pub fn set_margin(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin", margin.to_string());
        }

        pub fn set_margin_bottom(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-bottom", margin.to_string());
        }

        pub fn set_margin_left(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-left", margin.to_string());
        }

        pub fn set_margin_right(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-right", margin.to_string());
        }

        pub fn set_margin_top(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-top", margin.to_string());
        }
    };
}

#[macro_export]
macro_rules! paragraph {
    ($acc:ident) => {
        pub fn set_text_align_source(&mut self, align: TextAlignSource) {
            self.$acc()
                .set_attr("style:text-align-source", align.to_string());
        }

        pub fn set_text_align(&mut self, align: TextAlign) {
            self.$acc().set_attr("fo:text-align", align.to_string());
        }

        pub fn set_text_indent(&mut self, indent: Length) {
            self.$acc().set_attr("fo:text-indent", indent.to_string());
        }

        pub fn set_line_spacing(&mut self, spacing: Length) {
            self.$acc()
                .set_attr("style:line-spacing", spacing.to_string());
        }

        pub fn set_number_lines(&mut self, number: bool) {
            self.$acc()
                .set_attr("text:number-lines", number.to_string());
        }

        pub fn set_vertical_align_para(&mut self, align: ParaAlignVertical) {
            self.$acc()
                .set_attr("style:vertical-align", align.to_string());
        }
    };
}

#[macro_export]
macro_rules! text {
    ($acc:ident) => {
        pub fn set_color(&mut self, color: Rgb<u8>) {
            self.$acc().set_attr("fo:color", color_string(color));
        }

        pub fn set_font_name<S: Into<String>>(&mut self, name: S) {
            self.$acc().set_attr("style:font-name", name.into());
        }

        pub fn set_font_attr(&mut self, size: Length, bold: bool, italic: bool) {
            self.set_font_size(size);
            if bold {
                self.set_font_italic();
            }
            if italic {
                self.set_font_bold();
            }
        }

        pub fn set_font_size(&mut self, size: Length) {
            self.$acc().set_attr("fo:font-size", size.to_string());
        }

        pub fn set_font_size_percent(&mut self, size: f64) {
            self.$acc().set_attr("fo:font-size", percent_string(size));
        }

        pub fn set_font_italic(&mut self) {
            self.$acc().set_attr("fo:font-style", "italic".to_string());
        }

        pub fn set_font_style(&mut self, style: TextStyle) {
            self.$acc().set_attr("fo:font-style", style.to_string());
        }

        pub fn set_font_bold(&mut self) {
            self.$acc()
                .set_attr("fo:font-weight", TextWeight::Bold.to_string());
        }

        pub fn set_font_weight(&mut self, weight: TextWeight) {
            self.$acc().set_attr("fo:font-weight", weight.to_string());
        }

        pub fn set_letter_spacing(&mut self, spacing: Length) {
            self.$acc()
                .set_attr("fo:letter-spacing", spacing.to_string());
        }

        pub fn set_letter_spacing_normal(&mut self) {
            self.$acc()
                .set_attr("fo:letter-spacing", "normal".to_string());
        }

        pub fn set_text_shadow(
            &mut self,
            x_offset: Length,
            y_offset: Length,
            blur: Option<Length>,
            color: Rgb<u8>,
        ) {
            self.$acc().set_attr(
                "fo:text-shadow",
                shadow_string(x_offset, y_offset, blur, color),
            );
        }

        pub fn set_text_position(&mut self, pos: TextPosition) {
            self.$acc().set_attr("style:text-position", pos.to_string());
        }

        pub fn set_text_transform(&mut self, trans: TextTransform) {
            self.$acc().set_attr("fo:text-transform", trans.to_string());
        }

        pub fn set_font_relief(&mut self, relief: TextRelief) {
            self.$acc()
                .set_attr("style:font-relief", relief.to_string());
        }

        pub fn set_font_line_through_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("style:text-line-through-color", color_string(color));
        }

        pub fn set_font_line_through_style(&mut self, lstyle: LineStyle) {
            self.$acc()
                .set_attr("style:text-line-through-style", lstyle.to_string());
        }

        pub fn set_font_line_through_mode(&mut self, lmode: LineMode) {
            self.$acc()
                .set_attr("style:text-line-through-mode", lmode.to_string());
        }

        pub fn set_font_line_through_type(&mut self, ltype: LineType) {
            self.$acc()
                .set_attr("style:text-line-through-type", ltype.to_string());
        }

        pub fn set_font_line_through_text<S: Into<String>>(&mut self, text: S) {
            self.$acc()
                .set_attr("style:text-line-through-text", text.into());
        }

        pub fn set_font_line_through_text_style<S: Into<String>>(&mut self, style_ref: S) {
            self.$acc()
                .set_attr("style:text-line-through-text-style", style_ref.into());
        }

        pub fn set_font_line_through_width(&mut self, lwidth: LineWidth) {
            self.$acc()
                .set_attr("style:text-line-through-width", lwidth.to_string());
        }

        pub fn set_font_text_outline(&mut self, outline: bool) {
            self.$acc()
                .set_attr("style:text-outline", outline.to_string());
        }

        pub fn set_font_underline_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("style:text-underline-color", color_string(color));
        }

        pub fn set_font_underline_style(&mut self, lstyle: LineStyle) {
            self.$acc()
                .set_attr("style:text-underline-style", lstyle.to_string());
        }

        pub fn set_font_underline_type(&mut self, ltype: LineType) {
            self.$acc()
                .set_attr("style:text-underline-type", ltype.to_string());
        }

        pub fn set_font_underline_mode(&mut self, lmode: LineMode) {
            self.$acc()
                .set_attr("style:text-underline-mode", lmode.to_string());
        }

        pub fn set_font_underline_width(&mut self, lwidth: LineWidth) {
            self.$acc()
                .set_attr("style:text-underline-width", lwidth.to_string());
        }

        pub fn set_font_overline_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("style:text-overline-color", color_string(color));
        }

        pub fn set_font_overline_style(&mut self, lstyle: LineStyle) {
            self.$acc()
                .set_attr("style:text-overline-style", lstyle.to_string());
        }

        pub fn set_font_overline_type(&mut self, ltype: LineType) {
            self.$acc()
                .set_attr("style:text-overline-type", ltype.to_string());
        }

        pub fn set_font_overline_mode(&mut self, lmode: LineMode) {
            self.$acc()
                .set_attr("style:text-overline-mode", lmode.to_string());
        }

        pub fn set_font_overline_width(&mut self, lwidth: LineWidth) {
            self.$acc()
                .set_attr("style:text-overline-width", lwidth.to_string());
        }
    };
}
