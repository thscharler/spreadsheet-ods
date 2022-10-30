/// Generates the common features of all value formats.
macro_rules! valueformat {
    ($format:ident, $valuetype:expr) => {
        /// Formatting for Boolean.
        #[derive(Debug, Clone)]
        pub struct $format {
            /// Name
            name: String,
            /// Origin information.
            origin: StyleOrigin,
            /// Usage of this style.
            styleuse: StyleUse,
            /// Properties of the format.
            attr: AttrMap2,
            /// Cell text styles
            textstyle: AttrMap2,
            /// Parts of the format.
            parts: Vec<FormatPart>,
            /// Style map data.
            stylemaps: Option<Vec<StyleMap>>,
        }

        impl $format {
            /// New, empty.
            pub fn new_empty() -> Self {
                Self {
                    name: String::from(""),
                    origin: StyleOrigin::Styles,
                    styleuse: StyleUse::Default,
                    attr: Default::default(),
                    textstyle: Default::default(),
                    parts: Default::default(),
                    stylemaps: None,
                }
            }

            /// New, with name.
            pub fn new_named<S: Into<String>>(name: S) -> Self {
                Self {
                    name: name.into(),
                    origin: StyleOrigin::Styles,
                    styleuse: StyleUse::Default,
                    attr: Default::default(),
                    textstyle: Default::default(),
                    parts: Default::default(),
                    stylemaps: None,
                }
            }

            /// New, with name.
            pub fn new_localized<S: Into<String>>(name: S, locale: Locale) -> Self {
                let mut v = Self {
                    name: name.into(),
                    origin: StyleOrigin::Styles,
                    styleuse: StyleUse::Default,
                    attr: Default::default(),
                    textstyle: Default::default(),
                    parts: Default::default(),
                    stylemaps: None,
                };
                v.set_locale(locale);
                v
            }

            number_locale!(attr);
            number_title!(attr);
            number_transliteration_locale!(attr);
            number_transliteration_format!(attr);
            number_transliteration_style!(attr);
            style_display_name!(attr);
            style_volatile!(attr);

            fo_background_color!(textstyle);
            fo_color!(textstyle);
            // fo_locale!(textstyle);
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
            style_rotation_angle!(textstyle);
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

        impl ValueFormatTrait for $format {
            /// Returns a reference name for this value format.
            fn format_ref(&self) -> ValueFormatRef {
                ValueFormatRef::from(self.name().as_str())
            }

            /// The style:name attribute specifies names that reference style mechanisms.
            fn set_name<S: Into<String>>(&mut self, name: S) {
                self.name = name.into();
            }

            /// The style:name attribute specifies names that reference style mechanisms.
            fn name(&self) -> &String {
                &self.name
            }

            /// Returns the value type.
            fn value_type(&self) -> ValueType {
                $valuetype
            }

            /// Sets the storage location for this ValueFormat. Either content.xml
            /// or styles.xml.
            fn set_origin(&mut self, origin: StyleOrigin) {
                self.origin = origin;
            }

            /// Returns the storage location.
            fn origin(&self) -> StyleOrigin {
                self.origin
            }

            /// How is the style used in the document.
            fn set_styleuse(&mut self, styleuse: StyleUse) {
                self.styleuse = styleuse;
            }

            /// How is the style used in the document.
            fn styleuse(&self) -> StyleUse {
                self.styleuse
            }

            /// All direct attributes of the number:xxx-style tag.
            fn attrmap(&self) -> &AttrMap2 {
                &self.attr
            }

            /// All direct attributes of the number:xxx-style tag.
            fn attrmap_mut(&mut self) -> &mut AttrMap2 {
                &mut self.attr
            }

            /// Text style attributes.
            fn textstyle(&self) -> &AttrMap2 {
                &self.textstyle
            }

            /// Text style attributes.
            fn textstyle_mut(&mut self) -> &mut AttrMap2 {
                &mut self.textstyle
            }

            /// Adds a format part.
            fn push_part(&mut self, part: FormatPart) {
                self.parts.push(part);
            }

            /// Adds all format parts.
            fn push_parts(&mut self, partvec: &mut Vec<FormatPart>) {
                self.parts.append(partvec);
            }

            /// Returns the parts.
            fn parts(&self) -> &Vec<FormatPart> {
                &self.parts
            }

            /// Returns the mutable parts.
            fn parts_mut(&mut self) -> &mut Vec<FormatPart> {
                &mut self.parts
            }

            /// Adds a stylemap.
            fn push_stylemap(&mut self, stylemap: StyleMap) {
                self.stylemaps.get_or_insert_with(Vec::new).push(stylemap);
            }

            /// Returns the stylemaps
            fn stylemaps(&self) -> Option<&Vec<StyleMap>> {
                self.stylemaps.as_ref()
            }

            /// Returns the mutable stylemap.
            fn stylemaps_mut(&mut self) -> &mut Vec<StyleMap> {
                self.stylemaps.get_or_insert_with(Vec::new)
            }
        }
    };
}

macro_rules! part_number {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:number element specifies the display formatting properties for a decimal
        /// number.
        ///
        /// Can be used with ValueTypes:
        /// * Currency
        /// * Number
        /// * Percentage
        #[must_use]
        pub fn part_number(&mut self) -> PartNumberBuilder<'_, Self> {
            PartNumberBuilder::new(self)
        }
    };
}

macro_rules! part_fill_character {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:fill-character element specifies a Unicode character that is displayed
        /// repeatedly at the position where the element occurs. The character specified is repeated as many
        /// times as possible, but the total resulting string shall not exceed the given cell content area.
        ///
        /// Fill characters may not fill all the available space in a cell. The distribution of the
        /// remaining space is implementation-dependent.
        ///
        /// Can be used with ValueTypes:
        /// * Currency
        /// * DateTime
        /// * Number
        /// * Percentage
        /// * Text
        /// * TimeDuration
        #[must_use]
        pub fn part_fill_character(&mut self) -> PartFillCharacterBuilder<'_, Self> {
            PartFillCharacterBuilder::new(self)
        }
    };
}

macro_rules! part_scientific {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:scientific-number element specifies the display formatting properties for a
        /// number style that should be displayed in scientific format.
        ///
        /// Can be used with ValueTypes:
        /// * Number
        #[must_use]
        pub fn part_scientific(&mut self) -> PartScientificBuilder<'_, Self> {
            PartScientificBuilder::new(self)
        }
    };
}

macro_rules! part_fraction {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:fraction element specifies the display formatting properties for a number style
        /// that should be displayed as a fraction.
        ///
        /// Can be used with ValueTypes:
        /// * Number
        #[must_use]
        pub fn part_fraction(&mut self) -> PartFractionBuilder<'_, Self> {
            PartFractionBuilder::new(self)
        }
    };
}

macro_rules! part_currency {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:currency-symbol element specifies whether a currency symbol is displayed in
        /// a currency style.
        ///
        /// Can be used with ValueTypes:
        /// * Currency
        #[must_use]
        pub fn part_currency(&mut self) -> PartCurrencySymbolBuilder<'_, Self> {
            PartCurrencySymbolBuilder::new(self)
        }
    };
}

macro_rules! part_day {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:day element specifies a day of a month in a date.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_day(&mut self) -> PartDayBuilder<'_, Self> {
            PartDayBuilder::new(self)
        }
    };
}

macro_rules! part_month {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:month element specifies a month in a date.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_month(&mut self) -> PartMonthBuilder<'_, Self> {
            PartMonthBuilder::new(self)
        }
    };
}

macro_rules! part_year {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:year element specifies a year in a date
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_year(&mut self) -> PartYearBuilder<'_, Self> {
            PartYearBuilder::new(self)
        }
    };
}

macro_rules! part_era {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:era element specifies an era in which a year is counted
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_era(&mut self) -> PartEraBuilder<'_, Self> {
            PartEraBuilder::new(self)
        }
    };
}

macro_rules! part_day_of_week {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:day-of-week element specifies a day of a week in a date
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_day_of_week(&mut self) -> PartDayOfWeekBuilder<'_, Self> {
            PartDayOfWeekBuilder::new(self)
        }
    };
}

macro_rules! part_week_of_year {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:week-of-year element specifies a week of a year in a date.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_week_of_year(&mut self) -> PartWeekOfYearBuilder<'_, Self> {
            PartWeekOfYearBuilder::new(self)
        }
    };
}

macro_rules! part_quarter {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:quarter element specifies a quarter of the year in a date
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        #[must_use]
        pub fn part_quarter(&mut self) -> PartQuarterBuilder<'_, Self> {
            PartQuarterBuilder::new(self)
        }
    };
}

macro_rules! part_hours {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:hours element specifies whether hours are displayed as part of a date or time.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        /// * TimeDuration
        #[must_use]
        pub fn part_hours(&mut self) -> PartHoursBuilder<'_, Self> {
            PartHoursBuilder::new(self)
        }
    };
}

macro_rules! part_minutes {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:minutes element specifies whether minutes are displayed as part of a date or
        /// time.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        /// * TimeDuration
        #[must_use]
        pub fn part_minutes(&mut self) -> PartMinutesBuilder<'_, Self> {
            PartMinutesBuilder::new(self)
        }
    };
}

macro_rules! part_seconds {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:seconds element specifies whether seconds are displayed as part of a date or
        /// time.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        /// * TimeDuration
        #[must_use]
        pub fn part_seconds(&mut self) -> PartSecondsBuilder<'_, Self> {
            PartSecondsBuilder::new(self)
        }
    };
}

macro_rules! part_am_pm {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:am-pm element specifies whether AM/PM is included as part of a date or time.
        /// If a number:am-pm element is contained in a date or time style, hours are displayed using
        /// values from 1 to 12 only.
        ///
        /// Can be used with ValueTypes:
        /// * DateTime
        /// * TimeDuration
        #[must_use]
        pub fn part_am_pm(&mut self) -> PartAmPmBuilder<'_, Self> {
            PartAmPmBuilder::new(self)
        }
    };
}

macro_rules! part_boolean {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:boolean element marks the position of the Boolean value of a Boolean style.
        ///
        /// Can be used with ValueTypes:
        /// * Boolean
        #[must_use]
        pub fn part_boolean(&mut self) -> PartBooleanBuilder<'_, Self> {
            PartBooleanBuilder::new(self)
        }
    };
}

macro_rules! part_text {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:text element contains any fixed text for a data style.
        ///
        /// Can be used with ValueTypes:
        /// * Boolean
        /// * Currency
        /// * DateTime
        /// * Number
        /// * Percentage
        /// * Text
        /// * TimeDuration
        #[must_use]
        pub fn part_text<S: Into<String>>(&mut self, text: S) -> PartTextBuilder<'_, Self> {
            PartTextBuilder::new(self).text(text.into())
        }
    };
}

macro_rules! part_text_content {
    () => {
        /// Adds a format part to this format.
        ///
        /// The number:text-content element marks the position of variable text content of a text
        /// style.
        ///
        /// Can be used with ValueTypes:
        /// * Text
        #[must_use]
        pub fn part_text_content(&mut self) -> PartTextContentBuilder<'_, Self> {
            PartTextContentBuilder::new(self)
        }
    };
}

// macro_rules! style_xxx {
//     () => {};
// }
