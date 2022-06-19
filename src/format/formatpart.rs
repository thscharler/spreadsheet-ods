use crate::attrmap2::AttrMap2;
use chrono::{Duration, NaiveDateTime};
use icu_locid::Locale;
use std::fmt::{Display, Formatter};

/// Identifies the structural parts of a value format.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatPartType {
    Number,
    FillCharacter,
    ScientificNumber,
    Fraction,
    CurrencySymbol,
    Day,
    Month,
    Year,
    Era,
    DayOfWeek,
    WeekOfYear,
    Quarter,
    Hours,
    Minutes,
    Seconds,
    AmPm,
    Boolean,
    //EmbeddedText,
    Text,
    TextContent,
}

/// One structural part of a value format.
#[derive(Debug, Clone)]
pub struct FormatPart {
    /// What kind of format part is this?
    part_type: FormatPartType,
    /// Properties of this part.
    attr: AttrMap2,
    /// Some content.
    content: Option<String>,
}

/// Flag for several PartTypes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatNumberStyle {
    Short,
    Long,
}

impl Display for FormatNumberStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FormatNumberStyle::Short => write!(f, "short"),
            FormatNumberStyle::Long => write!(f, "long"),
        }
    }
}

/// Flag for several PartTypes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatTextual {
    Numeric,
    Textual,
}

impl Display for FormatTextual {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FormatTextual::Numeric => write!(f, "false"),
            FormatTextual::Textual => write!(f, "true"),
        }
    }
}

/// Flag for several PartTypes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatMonth {
    Nominativ,
    Possessiv,
}

impl Display for FormatMonth {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FormatMonth::Nominativ => write!(f, "false"),
            FormatMonth::Possessiv => write!(f, "true"),
        }
    }
}

/// Calendar types.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatCalendar {
    Default,
    Gregorian,
    Gengou,
    Roc,
    Hanja,
    Hijri,
    Jewish,
    Buddhist,
}

impl Display for FormatCalendar {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FormatCalendar::Gregorian => write!(f, "gregorian"),
            FormatCalendar::Gengou => write!(f, "gengou"),
            FormatCalendar::Roc => write!(f, "ROC"),
            FormatCalendar::Hanja => write!(f, "hanja"),
            FormatCalendar::Hijri => write!(f, "hijri"),
            FormatCalendar::Jewish => write!(f, "jewish"),
            FormatCalendar::Buddhist => write!(f, "buddhist"),
            FormatCalendar::Default => Ok(()),
        }
    }
}

impl FormatPart {
    /// New, empty
    pub fn new(ftype: FormatPartType) -> Self {
        FormatPart {
            part_type: ftype,
            attr: Default::default(),
            content: None,
        }
    }

    /// New, with string content.
    pub fn new_with_content<S: Into<String>>(ftype: FormatPartType, content: S) -> Self {
        FormatPart {
            part_type: ftype,
            attr: Default::default(),
            content: Some(content.into()),
        }
    }

    /// The <number:number> element specifies the display formatting properties for a decimal
    /// number.
    /// The <number:number> element is usable within the following elements:
    /// * <number:currencystyle> 16.29.8,
    /// * <number:number-style> 16.29.2 and
    /// * <number:percentage-style> 16.29.10.
    ///
    /// The <number:number> element has the following attributes:
    /// * number:decimal-places 19.343,
    /// * number:decimal-replacement 19.344
    /// * number:display-factor 19.346
    /// * number:grouping 19.350
    /// * number:min-decimal-places 19.356 and
    /// * number:mininteger-digits 19.355.
    ///
    /// The <number:number> element has the following child element: <number:embedded-text>
    /// 16.29.4.
    ///
    pub fn new_number(
        decimal_places: u8,
        grouping: bool,
        min_decimal_places: u8,
        mininteger_digits: u8,
        display_factor: Option<f64>,
        decimal_replacement: Option<char>,
    ) -> Self {
        let mut p = FormatPart::new(FormatPartType::Number);
        p.set_attr("number:min-integer-digits", 1.to_string());
        p.set_attr("number:decimal-places", decimal_places.to_string());
        if let Some(decimal_replacement) = decimal_replacement {
            p.set_attr(
                "number:decimal-replacement",
                decimal_replacement.to_string(),
            );
        }
        if let Some(display_factor) = display_factor {
            p.set_attr("number:display-factor", display_factor.to_string());
        }
        p.set_attr("number:mininteger-digits", mininteger_digits.to_string());
        p.set_attr("number:min-decimal-places", min_decimal_places.to_string());
        if grouping {
            p.set_attr("number:grouping", String::from("true"));
        }

        // TODO: number:embedded-text
        p
    }

    /// The <number:fill-character> element specifies a Unicode character that is displayed
    /// repeatedly at the position where the element occurs. The character specified is repeated as many
    /// times as possible, but the total resulting string shall not exceed the given cell content area.
    ///
    /// Fill characters may not fill all the available space in a cell. The distribution of the
    /// remaining space is implementation-dependent.
    ///
    /// The <number:fill-character> element is usable within the following elements:
    /// * <number:currency-style> 16.29.8,
    /// * <number:date-style> 16.29.11,
    /// * <number:number-style> 16.29.2,
    /// * <number:percentage-style> 16.29.10,
    /// * <number:text-style> 16.29.26 and
    /// * <number:time-style> 16.29.19.
    ///
    /// The <number:fill-character> element has no attributes.
    /// The <number:fill-character> element has no child elements.
    /// The <number:fill-character> element has character data content.
    pub fn new_fill_character(fill_character: char) -> Self {
        let mut p = FormatPart::new(FormatPartType::FillCharacter);
        p.set_content(fill_character.to_string());
        p
    }

    /// The <number:fraction> element specifies the display formatting properties for a number style
    /// that should be displayed as a fraction.
    ///
    /// The <number:fraction> element is usable within the following element:
    /// * <number:numberstyle> 16.29.2.
    ///
    /// The <number:fraction> element has the following attributes:
    /// * number:denominatorvalue 19.345,
    /// * number:grouping 19.350,
    /// * number:max-denominator-value 19.352,
    /// * number:min-denominator-digits 19.353,
    /// * number:min-integer-digits 19.355 and
    /// * number:min-numerator-digits 19.357.
    ///
    /// The <number:fraction> element has no child elements.
    pub fn new_fraction(
        denominatorvalue: u32,
        min_denominator_digits: u8,
        min_integer_digits: u8,
        min_numerator_digits: u8,
        grouping: bool,
        max_denominator_value: Option<u8>,
    ) -> Self {
        let mut p = Self::new(FormatPartType::Fraction);
        p.set_attr("number:denominator-value", denominatorvalue.to_string());
        if let Some(max_denominator_value) = max_denominator_value {
            p.set_attr(
                "number:max-denominator-value",
                max_denominator_value.to_string(),
            );
        }
        p.set_attr(
            "number:min-denominator-digits",
            min_denominator_digits.to_string(),
        );
        p.set_attr("number:min-integer-digits", min_integer_digits.to_string());
        p.set_attr(
            "number:min-numerator-digits",
            min_numerator_digits.to_string(),
        );
        if grouping {
            p.set_attr("number:grouping", String::from("true"));
        }
        p
    }

    /// The <number:scientific-number> element specifies the display formatting properties for a
    /// number style that should be displayed in scientific format.
    ///
    /// The <number:scientific-number> element is usable within the following element:
    /// * <number:number-style> 16.27.2.
    ///
    /// The <number:scientific-number> element has the following attributes:
    /// * number:decimal-places 19.343.4,
    /// * number:grouping 19.348,
    /// * number:min-exponentdigits 19.351 and
    /// * number:min-integer-digits 19.352.
    ///
    /// The <number:scientific-number> element has no child elements.
    pub fn new_scientific_number(
        decimal_places: u8,
        grouping: bool,
        min_exponentdigits: Option<u8>,
        min_integer_digits: Option<u8>,
    ) -> Self {
        let mut p = Self::new(FormatPartType::ScientificNumber);
        p.set_attr("number:decimal-places", decimal_places.to_string());
        if grouping {
            p.set_attr("number:grouping", String::from("true"));
        }
        if let Some(min_exponentdigits) = min_exponentdigits {
            p.set_attr("number:min-exponentdigits", min_exponentdigits.to_string());
        }
        if let Some(min_integer_digits) = min_integer_digits {
            p.set_attr("number:min-integer-digits", min_integer_digits.to_string());
        }
        p
    }

    /// The <number:currency-symbol> element specifies whether a currency symbol is displayed in
    /// a currency style.
    /// The content of this element is the text that is displayed as the currency symbol.
    /// If the element is empty or contains white space characters only, the default currency
    /// symbol for the currency style or the language and country of the currency style is displayed.
    ///
    /// The <number:currency-symbol> element is usable within the following element:
    /// * <number:currency-style> 16.27.7.
    ///
    /// The <number:currency-symbol> element has the following attributes:
    /// * number:country 19.342,
    /// * number:language 19.349,
    /// * number:rfc-language-tag 19.356 and
    /// * number:script 19.357.
    ///
    /// The <number:currency-symbol> element has no child elements.
    /// The <number:currency-symbol> element has character data content.
    pub fn new_currency_symbol<S>(locale: Locale, symbol: S) -> Self
    where
        S: Into<String>,
    {
        let mut p = Self::new_with_content(FormatPartType::CurrencySymbol, symbol);
        p.set_attr("number:language", locale.id.language.to_string());
        if let Some(region) = locale.id.region {
            p.set_attr("number:country", region.to_string());
        }
        p
    }

    /// The <number:day> element specifies a day of a month in a date.
    ///
    /// The <number:day> element is usable within the following element:
    /// * <number:date-style> 16.27.10.
    ///
    /// The <number:day> element has the following attributes:
    /// * number:calendar 19.341 and
    /// * number:style 19.358.2.
    ///
    /// The <number:day> element has no child elements.
    pub fn new_day(style: FormatNumberStyle, calendar: FormatCalendar) -> Self {
        let mut p = Self::new(FormatPartType::Day);
        p.set_attr("number:style", style.to_string());
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:month> element specifies a month in a date.
    /// The <number:month> element is usable within the following element:
    /// * <number:date-style> 16.27.10.
    /// The <number:month> element has the following attributes:
    /// number:calendar 19.341,
    /// number:possessive-form 19.355,
    /// number:style 19.358.7 and
    /// number:textual 19.359.
    ///
    /// The <number:month> element has no child elements
    pub fn new_month(
        style: FormatNumberStyle,
        textual: FormatTextual,
        possessive_form: FormatMonth,
        calendar: FormatCalendar,
    ) -> Self {
        let mut p = Self::new(FormatPartType::Month);
        p.set_attr("number:style", style.to_string());
        p.set_attr("number:textual", textual.to_string());
        if possessive_form != FormatMonth::Possessiv {
            p.set_attr("number:possessive-form", true.to_string());
        }
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:year> element specifies a year in a date.
    /// The <number:year> element is usable within the following element:
    /// * <number:date-style> 16.27.10.
    ///
    /// The <number:year> element has the following attributes:
    /// * number:calendar 19.341 and
    /// * number:style 19.358.10.
    ///
    /// The <number:year> element has no child elements.
    pub fn new_year(style: FormatNumberStyle, calendar: FormatCalendar) -> Self {
        let mut p = Self::new(FormatPartType::Year);
        p.set_attr("number:style", style.to_string());
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:era> element specifies an era in which a year is counted.
    ///
    /// The <number:era> element is usable within the following element:
    /// * <number:date-style> 16.27.10.
    ///
    /// The <number:era> element has the following attributes:
    /// * number:calendar 19.341 and
    /// * number:style 19.358.4.
    ///
    /// The <number:era> element has no child elements
    pub fn new_era(number: FormatNumberStyle, calendar: FormatCalendar) -> Self {
        let mut p = Self::new(FormatPartType::Era);
        p.set_attr("number:style", number.to_string());
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:day-of-week> element specifies a day of a week in a date.
    ///
    /// The <number:day-of-week> element is usable within the following element:
    /// * <number:datestyle> 16.27.10.
    ///
    /// The <number:day-of-week> element has the following attributes:
    /// * number:calendar 19.341 and
    /// * number:style 19.358.3.
    ///
    /// The <number:day-of-week> element has no child elements.
    pub fn new_day_of_week(style: FormatNumberStyle, calendar: FormatCalendar) -> Self {
        let mut p = Self::new(FormatPartType::DayOfWeek);
        p.set_attr("number:style", style.to_string());
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:week-of-year> element specifies a week of a year in a date.
    ///
    /// The <number:week-of-year> element is usable within the following element:
    /// * <number:date-style> 16.27.10.
    ///
    /// The <number:week-of-year> element has the following attribute:
    /// * number:calendar 19.341.
    ///
    /// The <number:week-of-year> element has no child elements.
    pub fn new_week_of_year(calendar: FormatCalendar) -> Self {
        let mut p = Self::new(FormatPartType::WeekOfYear);
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:quarter> element specifies a quarter of the year in a date.
    ///
    /// The <number:quarter> element is usable within the following element:
    /// * <number:datestyle> 16.27.10.
    ///
    /// The <number:quarter> element has the following attributes:
    /// * number:calendar 19.341 and
    /// * number:style 19.358.8.
    ///
    /// The <number:quarter> element has no child elements
    pub fn new_quarter(style: FormatNumberStyle, calendar: FormatCalendar) -> Self {
        let mut p = Self::new(FormatPartType::Quarter);
        p.set_attr("number:style", style.to_string());
        if calendar != FormatCalendar::Default {
            p.set_attr("number:calendar", calendar.to_string());
        }
        p
    }

    /// The <number:hours> element specifies whether hours are displayed as part of a date or time.
    ///
    /// The <number:hours> element is usable within the following elements:
    /// * <number:datestyle> 16.27.10 and
    /// * <number:time-style> 16.27.18.
    ///
    /// The <number:hours> element has the following attribute:
    /// * number:style 19.358.5.
    ///
    /// The <number:hours> element has no child elements.
    pub fn new_hours(style: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Hours);
        p.set_attr("number:style", style.to_string());
        p
    }

    /// The <number:minutes> element specifies whether minutes are displayed as part of a date or
    /// time.
    /// The <number:minutes> element is usable within the following elements:
    /// * <number:datestyle> 16.27.10 and
    /// * <number:time-style> 16.27.18.
    ///
    /// The <number:minutes> element has the following attribute:
    /// * number:style 19.358.6.
    ///
    /// The <number:minutes> element has no child elements.
    pub fn new_minutes(style: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Minutes);
        p.set_attr("number:style", style.to_string());
        p
    }

    /// The <number:seconds> element specifies whether seconds are displayed as part of a date or
    /// time.
    ///
    /// The <number:seconds> element is usable within the following elements:
    /// * <number:datestyle> 16.27.10 and
    /// * <number:time-style> 16.27.18.
    ///
    /// The <number:seconds> element has the following attributes:
    /// * number:decimal-places 19.343.3 and
    /// * number:style 19.358.9.
    ///
    /// The <number:seconds> element has no child elements.
    pub fn new_seconds(style: FormatNumberStyle, decimal_places: u8) -> Self {
        let mut p = Self::new(FormatPartType::Seconds);
        p.set_attr("number:style", style.to_string());
        p.set_attr("number:decimal-places", decimal_places.to_string());
        p
    }

    /// The <number:am-pm> element specifies whether AM/PM is included as part of a date or time.
    /// If a <number:am-pm> element is contained in a date or time style, hours are displayed using
    /// values from 1 to 12 only.
    ///
    /// The <number:am-pm> element is usable within the following elements:
    /// * <number:datestyle> 16.27.10 and
    /// * <number:time-style> 16.27.18.
    ///
    /// The <number:am-pm> element has no attributes.
    /// The <number:am-pm> element has no child elements.
    pub fn new_am_pm() -> Self {
        Self::new(FormatPartType::AmPm)
    }

    /// The <number:boolean> element marks the position of the Boolean value of a Boolean style.
    ///
    /// The <number:boolean> element is usable within the following element:
    /// * <number:booleanstyle> 16.29.24.
    ///
    /// The <number:boolean> element has no attributes.
    /// The <number:boolean> element has no child elements.
    pub fn new_boolean() -> Self {
        FormatPart::new(FormatPartType::Boolean)
    }

    /// The <number:text> element contains any fixed text for a data style.
    ///
    /// The <number:text> element is usable within the following elements:
    /// * <number:booleanstyle> 16.27.23,
    /// * <number:currency-style> 16.27.7,
    /// * <number:date-style> 16.27.10,
    /// * <number:number-style> 16.27.2,
    /// * <number:percentage-style> 16.27.9,
    /// * <number:text-style> 16.27.25 and
    /// * <number:time-style> 16.27.18.
    ///
    /// The <number:text> element has no attributes.
    /// The <number:text> element has no child elements.
    /// The <number:text> element has character data content
    pub fn new_text<S: Into<String>>(text: S) -> Self {
        Self::new_with_content(FormatPartType::Text, text)
    }

    /// The <number:text-content> element marks the position of variable text content of a text
    /// style.
    ///
    /// The <number:text-content> element is usable within the following element:
    /// * <number:text-style> 16.27.25.
    ///
    /// The <number:text-content> element has no attributes.
    /// The <number:text-content> element has no child elements.
    pub fn new_text_content() -> Self {
        Self::new(FormatPartType::TextContent)
    }

    // The <number:embedded-text> element specifies text that is displayed at one specific position
    // within a number.
    //
    // The <number:embedded-text> element is usable within the following element:
    // * <number:number> 16.27.3.
    //
    // The <number:embedded-text> element has the following attribute:
    // * number:position 19.354.
    //
    // The <number:embedded-text> element has no child elements.
    // The <number:embedded-text> element has character data content.
    // pub fn new_embedded_text<S: Into<String>>(position: u8, text: S) -> Self {
    //     let mut p = Self::new(FormatPartType::EmbeddedText);
    //     p.set_attr("number:position", position.to_string());
    //     p.set_content(text);
    //     p
    // }

    /// Sets the kind of the part.
    pub fn set_part_type(&mut self, p_type: FormatPartType) {
        self.part_type = p_type;
    }

    /// What kind of part?
    pub fn part_type(&self) -> FormatPartType {
        self.part_type
    }

    /// General attributes.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Adds an attribute.
    pub fn set_attr(&mut self, name: &str, value: String) {
        self.attr.set_attr(name, value);
    }

    /// Returns a property or a default.
    pub fn attr_def<'a0, 'a1, S0, S1>(&'a1 self, name: S0, default: S1) -> &'a1 str
    where
        S0: Into<&'a0 str>,
        S1: Into<&'a1 str>,
    {
        if let Some(v) = self.attr.attr(name.into()) {
            v
        } else {
            default.into()
        }
    }

    /// Sets a textual content for this part. This is only used
    /// for text and currency-symbol.
    pub fn set_content<S: Into<String>>(&mut self, content: S) {
        self.content = Some(content.into());
    }

    /// Returns the text content.
    pub fn content(&self) -> Option<&String> {
        self.content.as_ref()
    }

    /// Tries to format the given boolean, and appends the result to buf.
    /// If this part does'nt match does nothing
    pub(crate) fn format_boolean(&self, buf: &mut String, b: bool) {
        match self.part_type {
            FormatPartType::Boolean => {
                buf.push_str(if b { "true" } else { "false" });
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given float, and appends the result to buf.
    /// If this part does'nt match does nothing
    pub(crate) fn format_float(&self, buf: &mut String, f: f64) {
        match self.part_type {
            FormatPartType::Number => {
                let dec = self.attr_def("number:decimal-places", "0").parse::<usize>();
                if let Ok(dec) = dec {
                    buf.push_str(&format!("{:.*}", dec, f));
                }
            }
            FormatPartType::ScientificNumber => {
                buf.push_str(&format!("{:e}", f));
            }
            FormatPartType::CurrencySymbol => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given string, and appends the result to buf.
    /// If this part does'nt match does nothing
    pub(crate) fn format_str(&self, buf: &mut String, s: &str) {
        match self.part_type {
            FormatPartType::TextContent => {
                buf.push_str(s);
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given DateTime, and appends the result to buf.
    /// Uses chrono::strftime for the implementation.
    /// If this part does'nt match does nothing
    #[allow(clippy::collapsible_else_if)]
    pub(crate) fn format_datetime(&self, buf: &mut String, d: &NaiveDateTime, h12: bool) {
        match self.part_type {
            FormatPartType::Day => {
                let is_long = self.attr_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%d").to_string());
                } else {
                    buf.push_str(&d.format("%-d").to_string());
                }
            }
            FormatPartType::Month => {
                let is_long = self.attr_def("number:style", "") == "long";
                let is_text = self.attr_def("number:textual", "") == "true";
                if is_text {
                    if is_long {
                        buf.push_str(&d.format("%b").to_string());
                    } else {
                        buf.push_str(&d.format("%B").to_string());
                    }
                } else {
                    if is_long {
                        buf.push_str(&d.format("%m").to_string());
                    } else {
                        buf.push_str(&d.format("%-m").to_string());
                    }
                }
            }
            FormatPartType::Year => {
                let is_long = self.attr_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%Y").to_string());
                } else {
                    buf.push_str(&d.format("%y").to_string());
                }
            }
            FormatPartType::DayOfWeek => {
                let is_long = self.attr_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%A").to_string());
                } else {
                    buf.push_str(&d.format("%a").to_string());
                }
            }
            FormatPartType::WeekOfYear => {
                let is_long = self.attr_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%W").to_string());
                } else {
                    buf.push_str(&d.format("%-W").to_string());
                }
            }
            FormatPartType::Hours => {
                let is_long = self.attr_def("number:style", "") == "long";
                if !h12 {
                    if is_long {
                        buf.push_str(&d.format("%H").to_string());
                    } else {
                        buf.push_str(&d.format("%-H").to_string());
                    }
                } else {
                    if is_long {
                        buf.push_str(&d.format("%I").to_string());
                    } else {
                        buf.push_str(&d.format("%-I").to_string());
                    }
                }
            }
            FormatPartType::Minutes => {
                let is_long = self.attr_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%M").to_string());
                } else {
                    buf.push_str(&d.format("%-M").to_string());
                }
            }
            FormatPartType::Seconds => {
                let is_long = self.attr_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%S").to_string());
                } else {
                    buf.push_str(&d.format("%-S").to_string());
                }
            }
            FormatPartType::AmPm => {
                buf.push_str(&d.format("%p").to_string());
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given Duration, and appends the result to buf.
    /// If this part does'nt match does nothing
    pub(crate) fn format_time_duration(&self, buf: &mut String, d: &Duration) {
        match self.part_type {
            FormatPartType::Hours => {
                buf.push_str(&d.num_hours().to_string());
            }
            FormatPartType::Minutes => {
                buf.push_str(&(d.num_minutes() % 60).to_string());
            }
            FormatPartType::Seconds => {
                buf.push_str(&(d.num_seconds() % 60).to_string());
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }
}
