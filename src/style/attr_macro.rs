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
