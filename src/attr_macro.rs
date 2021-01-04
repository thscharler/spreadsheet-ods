macro_rules! fo_background_color {
    ($acc:ident) => {
        /// Background-color
        pub fn set_background_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("fo:background-color", color_string(color));
        }
    };
}

macro_rules! fo_break {
    ($acc:ident) => {
        /// Pagebreak before
        pub fn set_break_before(&mut self, pagebreak: PageBreak) {
            self.$acc()
                .set_attr("fo:break-before", pagebreak.to_string());
        }

        /// Pagebreak after
        pub fn set_break_after(&mut self, pagebreak: PageBreak) {
            self.$acc()
                .set_attr("fo:break-after", pagebreak.to_string());
        }
    };
}

macro_rules! fo_keep_with_next {
    ($acc:ident) => {
        /// Keep with next
        pub fn set_keep_with_next(&mut self, keep_with_next: TextKeep) {
            self.$acc()
                .set_attr("fo:keep-with-next", keep_with_next.to_string());
        }
    };
}

macro_rules! style_shadow {
    ($acc:ident) => {
        /// Shadow
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

macro_rules! style_writing_mode {
    ($acc:ident) => {
        /// Writing-mode
        pub fn set_writing_mode(&mut self, writing_mode: WritingMode) {
            self.$acc()
                .set_attr("style:writing-mode", writing_mode.to_string());
        }
    };
}

macro_rules! fo_keep_together {
    ($acc:ident) => {
        /// Keep-together
        pub fn set_keep_together(&mut self, keep_together: TextKeep) {
            self.$acc()
                .set_attr("fo:keep-together", keep_together.to_string());
        }
    };
}

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

macro_rules! fo_padding {
    ($acc:ident) => {
        /// Padding for all sides.
        pub fn set_padding(&mut self, padding: Length) {
            self.$acc().set_attr("fo:padding", padding.to_string());
        }

        /// Padding
        pub fn set_padding_bottom(&mut self, padding: Length) {
            self.$acc()
                .set_attr("fo:padding-bottom", padding.to_string());
        }

        /// Padding
        pub fn set_padding_left(&mut self, padding: Length) {
            self.$acc().set_attr("fo:padding-left", padding.to_string());
        }

        /// Padding
        pub fn set_padding_right(&mut self, padding: Length) {
            self.$acc()
                .set_attr("fo:padding-right", padding.to_string());
        }

        /// Padding
        pub fn set_padding_top(&mut self, padding: Length) {
            self.$acc().set_attr("fo:padding-top", padding.to_string());
        }
    };
}

macro_rules! fo_margin {
    ($acc:ident) => {
        /// Margin for all sides.
        pub fn set_margin(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin", margin.to_string());
        }

        /// Margin
        pub fn set_margin_bottom(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-bottom", margin.to_string());
        }

        /// Margin
        pub fn set_margin_left(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-left", margin.to_string());
        }

        /// Margin
        pub fn set_margin_right(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-right", margin.to_string());
        }

        /// Margin
        pub fn set_margin_top(&mut self, margin: Length) {
            self.$acc().set_attr("fo:margin-top", margin.to_string());
        }
    };
}

// missing
// fo:hyphenation-keep 20.196,
// fo:hyphenation-ladder-count 20.197,
// fo:line-height 20.204,
// fo:orphans 20.214,
// fo:textalign-last 20.224,
// fo:widows 20.228,
// style:auto-textindent 20.246,
// style:background-transparency 20.247,
// style:contextual-spacing 20.255,
// style:font-independent-line-spacing 20.276,
// style:join-border 20.300,
// style:justify-single-word 20.301,
// style:line-break 20.315,
// style:line-height-at-least 20.317,
// style:page-number 20.328,
// style:punctuation-wrap 20.335,
// style:register-true 20.336,
// style:snap-to-layout-grid 20.361,
// style:tab-stop-distance 20.362,
// style:text-autospace 20.365,
// style:writing-modeautomatic 20.405,
// text:line-number 20.430
macro_rules! paragraph {
    ($acc:ident) => {
        /// Text alignment.
        pub fn set_text_align_source(&mut self, align: TextAlignSource) {
            self.$acc()
                .set_attr("style:text-align-source", align.to_string());
        }

        /// Text alignment.
        pub fn set_text_align(&mut self, align: TextAlign) {
            self.$acc().set_attr("fo:text-align", align.to_string());
        }

        /// Text indent.
        pub fn set_text_indent(&mut self, indent: Length) {
            self.$acc().set_attr("fo:text-indent", indent.to_string());
        }

        /// Line spacing.
        pub fn set_line_spacing(&mut self, spacing: Length) {
            self.$acc()
                .set_attr("style:line-spacing", spacing.to_string());
        }

        /// Line numbering.
        pub fn set_number_lines(&mut self, number: bool) {
            self.$acc()
                .set_attr("text:number-lines", number.to_string());
        }

        /// Vertical alignment for paragraphs.
        pub fn set_vertical_align_para(&mut self, align: ParaAlignVertical) {
            self.$acc()
                .set_attr("style:vertical-align", align.to_string());
        }
    };
}

// fo:backgroundcolor 20.182,
// fo:country 20.188,
// fo:font-family 20.189,
// fo:font-variant 20.192,
// fo:hyphenate 20.195,
// fo:hyphenation-push-char-count 20.198,
// fo:hyphenationremain-char-count 20.199,
// fo:language 20.202,
// fo:script 20.222,
// style:country-asian 20.256,
// style:country-complex 20.257,
// style:font-charset 20.268,
// style:font-charset-asian 20.269,
// style:font-charset-complex 20.270,
// style:font-family-asian 20.271,
// style:font-family-complex 20.272,
// style:font-family-generic 20.273,
// style:font-family-generic-asian 20.274,
// style:font-family-generic-complex 20.275,
// style:font-name-asian 20.278,
// style:font-name-complex 20.279,
// style:fontpitch 20.280,
// style:font-pitch-asian 20.281,
// style:font-pitch-complex 20.282,
// style:font-size-asian 20.284,
// style:font-sizecomplex 20.285,
// style:font-size-rel 20.286,
// style:font-size-rel-asian 20.287,
// style:font-size-rel-complex 20.288,
// style:font-style-asian 20.289,
// style:font-style-complex 20.290,
// style:font-style-name 20.291,
// style:fontstyle-name-asian 20.292,
// style:font-style-name-complex 20.293,
// style:fontweight-asian 20.294,
// style:font-weight-complex 20.295,
// style:language-asian 20.302,
// style:language-complex 20.303,
// style:letter-kerning 20.316,
// style:rfclanguage-tag 20.343,
// style:rfc-language-tag-asian 20.344,
// style:rfclanguage-tag-complex 20.345,
// style:script-asian 20.356,
// style:script-complex 20.357,
// style:script-type 20.358,
// style:text-blinking 20.366,
// style:textcombine 20.367,
// style:text-combine-end-char 20.369,
// style:text-combinestart-char 20.368,
// style:text-emphasize 20.370,
// style:text-position 20.384,
// style:text-rotation-angle 20.385,
// style:textrotation-scale 20.386,
// style:text-scale 20.387,
// style:use-window-font-color 20.395,
// text:condition 20.426,
// text:display 20.427.
macro_rules! text {
    ($acc:ident) => {
        /// Text color
        pub fn set_color(&mut self, color: Rgb<u8>) {
            self.$acc().set_attr("fo:color", color_string(color));
        }

        /// Text font.
        pub fn set_font_name<S: Into<String>>(&mut self, name: S) {
            self.$acc().set_attr("style:font-name", name.into());
        }

        /// Combined font attributes.
        pub fn set_font_attr(&mut self, size: Length, bold: bool, italic: bool) {
            self.set_font_size(size);
            if bold {
                self.set_font_italic();
            }
            if italic {
                self.set_font_bold();
            }
        }

        /// Font size.
        pub fn set_font_size(&mut self, size: Length) {
            self.$acc().set_attr("fo:font-size", size.to_string());
        }

        /// Font size as a percentage.
        pub fn set_font_size_percent(&mut self, size: f64) {
            self.$acc().set_attr("fo:font-size", percent_string(size));
        }

        /// Set to italic.
        pub fn set_font_italic(&mut self) {
            self.$acc().set_attr("fo:font-style", "italic".to_string());
        }

        /// Set font style.
        pub fn set_font_style(&mut self, style: FontStyle) {
            self.$acc().set_attr("fo:font-style", style.to_string());
        }

        /// Set to bold.
        pub fn set_font_bold(&mut self) {
            self.$acc()
                .set_attr("fo:font-weight", FontWeight::Bold.to_string());
        }

        /// Sets the font weight.
        pub fn set_font_weight(&mut self, weight: FontWeight) {
            self.$acc().set_attr("fo:font-weight", weight.to_string());
        }

        /// Sets the letter spacing.
        pub fn set_letter_spacing(&mut self, spacing: Length) {
            self.$acc()
                .set_attr("fo:letter-spacing", spacing.to_string());
        }

        /// Sets the letter spacing to normal.
        pub fn set_letter_spacing_normal(&mut self) {
            self.$acc()
                .set_attr("fo:letter-spacing", "normal".to_string());
        }

        /// Text shadow.
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

        /// Text positioning.        
        pub fn set_text_position(&mut self, pos: TextPosition) {
            self.$acc().set_attr("style:text-position", pos.to_string());
        }

        /// Transforms on the text.
        pub fn set_text_transform(&mut self, trans: TextTransform) {
            self.$acc().set_attr("fo:text-transform", trans.to_string());
        }

        /// Font style relief.
        pub fn set_font_relief(&mut self, relief: TextRelief) {
            self.$acc()
                .set_attr("style:font-relief", relief.to_string());
        }

        /// Color
        pub fn set_font_line_through_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("style:text-line-through-color", color_string(color));
        }

        /// Line through
        pub fn set_font_line_through_style(&mut self, lstyle: LineStyle) {
            self.$acc()
                .set_attr("style:text-line-through-style", lstyle.to_string());
        }

        /// Line through
        pub fn set_font_line_through_mode(&mut self, lmode: LineMode) {
            self.$acc()
                .set_attr("style:text-line-through-mode", lmode.to_string());
        }

        /// Line through
        pub fn set_font_line_through_type(&mut self, ltype: LineType) {
            self.$acc()
                .set_attr("style:text-line-through-type", ltype.to_string());
        }

        /// Line through
        pub fn set_font_line_through_text<S: Into<String>>(&mut self, text: S) {
            self.$acc()
                .set_attr("style:text-line-through-text", text.into());
        }

        /// References a text-style.
        pub fn set_font_line_through_text_style(&mut self, style_ref: TextStyleRef) {
            self.$acc()
                .set_attr("style:text-line-through-text-style", style_ref.to_string());
        }

        /// Line through
        pub fn set_font_line_through_width(&mut self, lwidth: LineWidth) {
            self.$acc()
                .set_attr("style:text-line-through-width", lwidth.to_string());
        }

        /// Outline
        pub fn set_font_text_outline(&mut self, outline: bool) {
            self.$acc()
                .set_attr("style:text-outline", outline.to_string());
        }

        /// Underlining
        pub fn set_font_underline_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("style:text-underline-color", color_string(color));
        }

        /// Underlining
        pub fn set_font_underline_style(&mut self, lstyle: LineStyle) {
            self.$acc()
                .set_attr("style:text-underline-style", lstyle.to_string());
        }

        /// Underlining
        pub fn set_font_underline_type(&mut self, ltype: LineType) {
            self.$acc()
                .set_attr("style:text-underline-type", ltype.to_string());
        }

        /// Underlining
        pub fn set_font_underline_mode(&mut self, lmode: LineMode) {
            self.$acc()
                .set_attr("style:text-underline-mode", lmode.to_string());
        }

        /// Underlining
        pub fn set_font_underline_width(&mut self, lwidth: LineWidth) {
            self.$acc()
                .set_attr("style:text-underline-width", lwidth.to_string());
        }

        /// Overlining
        pub fn set_font_overline_color(&mut self, color: Rgb<u8>) {
            self.$acc()
                .set_attr("style:text-overline-color", color_string(color));
        }

        /// Overlining
        pub fn set_font_overline_style(&mut self, lstyle: LineStyle) {
            self.$acc()
                .set_attr("style:text-overline-style", lstyle.to_string());
        }

        /// Overlining
        pub fn set_font_overline_type(&mut self, ltype: LineType) {
            self.$acc()
                .set_attr("style:text-overline-type", ltype.to_string());
        }

        /// Overlining
        pub fn set_font_overline_mode(&mut self, lmode: LineMode) {
            self.$acc()
                .set_attr("style:text-overline-mode", lmode.to_string());
        }

        /// Overlining
        pub fn set_font_overline_width(&mut self, lwidth: LineWidth) {
            self.$acc()
                .set_attr("style:text-overline-width", lwidth.to_string());
        }
    };
}

macro_rules! font_decl {
    ($acc:ident) => {
        /// External font family name.
        pub fn set_font_family<S: Into<String>>(&mut self, name: S) {
            self.$acc().set_attr("svg:font-family", name.into());
        }

        /// System generic name.
        pub fn set_font_family_generic<S: Into<String>>(&mut self, name: S) {
            self.$acc()
                .set_attr("style:font-family-generic", name.into());
        }

        /// Font pitch.
        pub fn set_font_pitch(&mut self, pitch: FontPitch) {
            self.$acc().set_attr("style:font-pitch", pitch.to_string());
        }
    };
}

macro_rules! svg_height {
    ($acc:ident) => {
        /// Height.
        pub fn set_height(&mut self, height: Length) {
            self.$acc().set_attr("svg:height", height.to_string());
        }
    };
}

macro_rules! fo_min_height {
    ($acc:ident) => {
        /// Minimum height.
        pub fn set_min_height(&mut self, height: Length) {
            self.$acc().set_attr("fo:min-height", height.to_string());
        }

        /// Minimum height as percentage.
        pub fn set_min_height_percent(&mut self, height: f64) {
            self.$acc()
                .set_attr("fo:min-height", percent_string(height));
        }
    };
}

macro_rules! style_dynamic_spacing {
    ($acc:ident) => {
        /// Dynamic spacing
        pub fn set_dynamic_spacing(&mut self, dynamic: bool) {
            self.$acc()
                .set_attr("style:dynamic-spacing", dynamic.to_string());
        }
    };
}
