//!
//! Defines ValueFormat for formatting related issues
//!
//! ```
//! use spreadsheet_ods::{ValueFormat, ValueType};
//! use spreadsheet_ods::format::FormatNumberStyle;
//!
//! let mut v = ValueFormat::new_with_name("dt0", ValueType::DateTime);
//! v.push_day(FormatNumberStyle::Long);
//! v.push_text(".");
//! v.push_month(FormatNumberStyle::Long);
//! v.push_text(".");
//! v.push_year(FormatNumberStyle::Long);
//! v.push_text(" ");
//! v.push_hours(FormatNumberStyle::Long);
//! v.push_text(":");
//! v.push_minutes(FormatNumberStyle::Long);
//! v.push_text(":");
//! v.push_seconds(FormatNumberStyle::Long);
//!
//! let mut v = ValueFormat::new_with_name("n3", ValueType::Number);
//! v.push_number(3, false);
//! ```
//! The output formatting is a rough approximation with the possibilities
//! offered by format! and chrono::format. Especially there is no trace of
//! i18n. But on the other hand the formatting rules are applied by LibreOffice
//! when opening the spreadsheet so typically nobody notices this.
//!

pub use crate::attrmap2::{AttrMap2, AttrMap2Trait};

use crate::style::stylemap::StyleMap;
use crate::style::units::{
    FontStyle, FontWeight, Length, LineMode, LineStyle, LineType, LineWidth, TextPosition,
    TextRelief, TextTransform,
};
use crate::style::{color_string, percent_string, shadow_string, StyleOrigin, StyleUse};
use crate::ValueType;
use chrono::NaiveDateTime;
use color::Rgb;
use std::fmt::{Display, Formatter};
use time::Duration;

#[derive(Debug)]
pub enum ValueFormatError {
    Format(String),
    NaN,
}

impl Display for ValueFormatError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            ValueFormatError::Format(s) => write!(f, "{}", s)?,
            ValueFormatError::NaN => write!(f, "Digit expected")?,
        }
        Ok(())
    }
}

impl std::error::Error for ValueFormatError {}

/// Actual textual formatting of values.
#[derive(Debug, Clone, Default)]
pub struct ValueFormat {
    /// Name
    name: String,
    country: Option<String>,
    language: Option<String>,
    script: Option<String>,
    /// Value type
    v_type: ValueType,
    /// Origin information.
    origin: StyleOrigin,
    /// Usage of this style.
    styleuse: StyleUse,
    /// Properties of the format.
    attr: AttrMap2,
    /// Cell text styles
    text_style: AttrMap2,
    /// Parts of the format.
    parts: Vec<FormatPart>,
    /// Style map data.
    stylemaps: Option<Vec<StyleMap>>,
}

impl ValueFormat {
    /// New, empty.
    pub fn new() -> Self {
        ValueFormat {
            name: String::from(""),
            country: None,
            language: None,
            script: None,
            v_type: ValueType::Text,
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            text_style: Default::default(),
            parts: Default::default(),
            stylemaps: None,
        }
    }

    /// New, with name.
    pub fn new_with_name<S: Into<String>>(name: S, value_type: ValueType) -> Self {
        ValueFormat {
            name: name.into(),
            country: None,
            language: None,
            script: None,
            v_type: value_type,
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            text_style: Default::default(),
            parts: Default::default(),
            stylemaps: None,
        }
    }

    /// Sets the name.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets the country.
    pub fn set_country<S: Into<String>>(&mut self, country: S) {
        self.country = Some(country.into());
    }

    /// Country
    pub fn country(&self) -> Option<&String> {
        self.country.as_ref()
    }

    /// Sets the language.
    pub fn set_language<S: Into<String>>(&mut self, language: S) {
        self.language = Some(language.into());
    }

    /// Language
    pub fn language(&self) -> Option<&String> {
        self.language.as_ref()
    }

    /// Sets the Script.
    pub fn set_script<S: Into<String>>(&mut self, script: S) {
        self.script = Some(script.into());
    }

    /// Script
    pub fn script(&self) -> Option<&String> {
        self.script.as_ref()
    }

    /// Sets the value type.
    pub fn set_value_type(&mut self, value_type: ValueType) {
        self.v_type = value_type;
    }

    /// Returns the value type.
    pub fn value_type(&self) -> ValueType {
        self.v_type
    }

    /// Sets the origin.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Returns the origin.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Style usage.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Returns the usage.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Text style attributes.
    pub fn text_style(&self) -> &AttrMap2 {
        &self.text_style
    }

    /// Text style attributes.
    pub fn text_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.text_style
    }

    text!(text_style_mut);

    /// Appends a format part.
    pub fn push_boolean(&mut self) {
        self.push_part(FormatPart::new_boolean());
    }

    /// Appends a format part.
    pub fn push_number(&mut self, decimal: u8, grouping: bool) {
        self.push_part(FormatPart::new_number(decimal, grouping));
    }

    /// Appends a format part.
    pub fn push_number_fix(&mut self, decimal: u8, grouping: bool) {
        self.push_part(FormatPart::new_number_fix(decimal, grouping));
    }

    /// Appends a format part.
    pub fn push_fraction(
        &mut self,
        denominator: u32,
        min_den_digits: u8,
        min_int_digits: u8,
        min_num_digits: u8,
        grouping: bool,
    ) {
        self.push_part(FormatPart::new_fraction(
            denominator,
            min_den_digits,
            min_int_digits,
            min_num_digits,
            grouping,
        ));
    }

    /// Appends a format part.
    pub fn push_scientific(&mut self, dec_places: u8) {
        self.push_part(FormatPart::new_scientific(dec_places));
    }

    /// Appends a format part.
    pub fn push_currency<S1, S2, S3>(&mut self, country: S1, language: S2, symbol: S3)
    where
        S1: Into<String>,
        S2: Into<String>,
        S3: Into<String>,
    {
        self.push_part(FormatPart::new_currency(country, language, symbol));
    }

    /// Appends a format part.
    pub fn push_day(&mut self, number: FormatNumberStyle) {
        self.push_part(FormatPart::new_day(number));
    }

    /// Appends a format part.
    pub fn push_month(&mut self, number: FormatNumberStyle) {
        self.push_part(FormatPart::new_month(number));
    }

    /// Appends a format part.
    pub fn push_year(&mut self, number: FormatNumberStyle) {
        self.push_part(FormatPart::new_year(number));
    }

    /// Appends a format part.
    pub fn push_era(&mut self, number: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.push_part(FormatPart::new_era(number, calendar));
    }

    /// Appends a format part.
    pub fn push_day_of_week(&mut self, number: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.push_part(FormatPart::new_day_of_week(number, calendar));
    }

    /// Appends a format part.
    pub fn push_week_of_year(&mut self, calendar: FormatCalendarStyle) {
        self.push_part(FormatPart::new_week_of_year(calendar));
    }

    /// Appends a format part.
    pub fn push_quarter(&mut self, number: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.push_part(FormatPart::new_quarter(number, calendar));
    }

    /// Appends a format part.
    pub fn push_hours(&mut self, number: FormatNumberStyle) {
        self.push_part(FormatPart::new_hours(number));
    }

    /// Appends a format part.
    pub fn push_minutes(&mut self, number: FormatNumberStyle) {
        self.push_part(FormatPart::new_minutes(number));
    }

    /// Appends a format part.
    pub fn push_seconds(&mut self, number: FormatNumberStyle) {
        self.push_part(FormatPart::new_seconds(number));
    }

    /// Appends a format part.
    pub fn push_am_pm(&mut self) {
        self.push_part(FormatPart::new_am_pm());
    }

    /// Appends a format part.
    pub fn push_embedded_text(&mut self, position: u8) {
        self.push_part(FormatPart::new_embedded_text(position));
    }

    /// Appends a format part.
    pub fn push_text<S: Into<String>>(&mut self, text: S) {
        self.push_part(FormatPart::new_text(text));
    }

    /// Appends a format part.
    pub fn push_text_content(&mut self) {
        self.push_part(FormatPart::new_text_content());
    }

    /// Adds a format part.
    pub fn push_part(&mut self, part: FormatPart) {
        self.parts.push(part);
    }

    /// Adds all format parts.
    #[allow(clippy::collapsible_if)]
    pub fn push_parts(&mut self, partvec: &mut Vec<FormatPart>) {
        self.parts.append(partvec);
    }

    /// Returns the parts.
    pub fn parts(&self) -> &Vec<FormatPart> {
        &self.parts
    }

    /// Returns the mutable parts.
    pub fn parts_mut(&mut self) -> &mut Vec<FormatPart> {
        &mut self.parts
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

    /// Tries to format.
    /// If there are no matching parts, does nothing.
    pub fn format_boolean(&self, b: bool) -> String {
        let mut buf = String::new();
        for p in &self.parts {
            p.format_boolean(&mut buf, b);
        }
        buf
    }

    /// Tries to format.
    /// If there are no matching parts, does nothing.
    pub fn format_float(&self, f: f64) -> String {
        let mut buf = String::new();
        for p in &self.parts {
            p.format_float(&mut buf, f);
        }
        buf
    }

    /// Tries to format.
    /// If there are no matching parts, does nothing.
    pub fn format_str<'a, S: Into<&'a str>>(&self, s: S) -> String {
        let mut buf = String::new();
        let s = s.into();
        for p in &self.parts {
            p.format_str(&mut buf, s);
        }
        buf
    }

    /// Tries to format.
    /// If there are no matching parts, does nothing.
    /// Should work reasonably. Don't ask me about other calenders.
    pub fn format_datetime(&self, d: &NaiveDateTime) -> String {
        let mut buf = String::new();

        let h12 = self
            .parts
            .iter()
            .any(|v| v.part_type == FormatPartType::AmPm);

        for p in &self.parts {
            p.format_datetime(&mut buf, d, h12);
        }
        buf
    }

    /// Tries to format. Should work reasonably.
    /// If there are no matching parts, does nothing.
    pub fn format_time_duration(&self, d: &Duration) -> String {
        let mut buf = String::new();
        for p in &self.parts {
            p.format_time_duration(&mut buf, d);
        }
        buf
    }
}

impl AttrMap2Trait for ValueFormat {
    fn attr_map(&self) -> &AttrMap2 {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }
}

/// Identifies the structural parts of a value format.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum FormatPartType {
    Boolean,
    Number,
    Fraction,
    Scientific,
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
    EmbeddedText,
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

impl AttrMap2Trait for FormatPart {
    fn attr_map(&self) -> &AttrMap2 {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }
}

/// Flag for several PartTypes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
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

/// Calendar types.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum FormatCalendarStyle {
    Gregorian,
    Gengou,
    ROC,
    Hanja,
    Hijri,
    Jewish,
    Buddhist,
}

impl Display for FormatCalendarStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FormatCalendarStyle::Gregorian => write!(f, "gregorian"),
            FormatCalendarStyle::Gengou => write!(f, "gengou"),
            FormatCalendarStyle::ROC => write!(f, "ROC"),
            FormatCalendarStyle::Hanja => write!(f, "hanja"),
            FormatCalendarStyle::Hijri => write!(f, "hijri"),
            FormatCalendarStyle::Jewish => write!(f, "jewish"),
            FormatCalendarStyle::Buddhist => write!(f, "buddhist"),
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

    /// Boolean Part
    pub fn new_boolean() -> Self {
        FormatPart::new(FormatPartType::Boolean)
    }

    /// Number format part.
    pub fn new_number(decimal: u8, grouping: bool) -> Self {
        let mut p = FormatPart::new(FormatPartType::Number);
        p.set_attr("number:min-integer-digits", 1.to_string());
        p.set_attr("number:decimal-places", decimal.to_string());
        p.set_attr("loext:min-decimal-places", 0.to_string());
        if grouping {
            p.set_attr("number:grouping", String::from("true"));
        }
        p
    }

    /// Number format part with fixed decimal places.
    pub fn new_number_fix(decimal: u8, grouping: bool) -> Self {
        let mut p = Self::new(FormatPartType::Number);
        p.set_attr("number:min-integer-digits", 1.to_string());
        p.set_attr("number:decimal-places", decimal.to_string());
        p.set_attr("loext:min-decimal-places", decimal.to_string());
        if grouping {
            p.set_attr("number:grouping", String::from("true"));
        }
        p
    }

    /// Format as a fraction.
    pub fn new_fraction(
        denominator: u32,
        min_den_digits: u8,
        min_int_digits: u8,
        min_num_digits: u8,
        grouping: bool,
    ) -> Self {
        let mut p = Self::new(FormatPartType::Fraction);
        p.set_attr("number:denominator-value", denominator.to_string());
        p.set_attr("number:min-denominator-digits", min_den_digits.to_string());
        p.set_attr("number:min-integer-digits", min_int_digits.to_string());
        p.set_attr("number:min-numerator-digits", min_num_digits.to_string());
        if grouping {
            p.set_attr("number:grouping", String::from("true"));
        }
        p
    }

    /// Format with scientific notation.
    pub fn new_scientific(dec_places: u8) -> Self {
        let mut p = Self::new(FormatPartType::Scientific);
        p.set_attr("number:decimal-places", dec_places.to_string());
        p
    }

    /// Currency symbol.
    pub fn new_currency<S1, S2, S3>(country: S1, language: S2, symbol: S3) -> Self
    where
        S1: Into<String>,
        S2: Into<String>,
        S3: Into<String>,
    {
        let mut p = Self::new_with_content(FormatPartType::CurrencySymbol, symbol);
        p.set_attr("number:country", country.into());
        p.set_attr("number:language", language.into());
        p
    }

    pub fn new_day(number: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Day);
        p.set_attr("number:style", number.to_string());
        p
    }

    pub fn new_month(number: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Month);
        p.set_attr("number:style", number.to_string());
        p
    }

    pub fn new_year(number: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Year);
        p.set_attr("number:style", number.to_string());
        p
    }

    pub fn new_era(number: FormatNumberStyle, calendar: FormatCalendarStyle) -> Self {
        let mut p = Self::new(FormatPartType::Era);
        p.set_attr("number:style", number.to_string());
        p.set_attr("number:calendar", calendar.to_string());
        p
    }

    pub fn new_day_of_week(number: FormatNumberStyle, calendar: FormatCalendarStyle) -> Self {
        let mut p = Self::new(FormatPartType::DayOfWeek);
        p.set_attr("number:style", number.to_string());
        p.set_attr("number:calendar", calendar.to_string());
        p
    }

    pub fn new_week_of_year(calendar: FormatCalendarStyle) -> Self {
        let mut p = Self::new(FormatPartType::WeekOfYear);
        p.set_attr("number:calendar", calendar.to_string());
        p
    }

    pub fn new_quarter(number: FormatNumberStyle, calendar: FormatCalendarStyle) -> Self {
        let mut p = Self::new(FormatPartType::Quarter);
        p.set_attr("number:style", number.to_string());
        p.set_attr("number:calendar", calendar.to_string());
        p
    }

    pub fn new_hours(number: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Hours);
        p.set_attr("number:style", number.to_string());
        p
    }

    pub fn new_minutes(number: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Minutes);
        p.set_attr("number:style", number.to_string());
        p
    }

    pub fn new_seconds(number: FormatNumberStyle) -> Self {
        let mut p = Self::new(FormatPartType::Seconds);
        p.set_attr("number:style", number.to_string());
        p
    }

    pub fn new_am_pm() -> Self {
        Self::new(FormatPartType::AmPm)
    }

    /// Whatever this is for ...
    pub fn new_embedded_text(position: u8) -> Self {
        let mut p = Self::new(FormatPartType::EmbeddedText);
        p.set_attr("number:position", position.to_string());
        p
    }

    /// Part with fixed text.
    pub fn new_text<S: Into<String>>(text: S) -> Self {
        Self::new_with_content(FormatPartType::Text, text)
    }

    /// Whatever this is for ...
    pub fn new_text_content() -> Self {
        Self::new(FormatPartType::TextContent)
    }

    /// Sets the kind of the part.
    pub fn set_part_type(&mut self, p_type: FormatPartType) {
        self.part_type = p_type;
    }

    /// What kind of part?
    pub fn part_type(&self) -> FormatPartType {
        self.part_type
    }

    /// Returns a property or a default.
    pub fn attr_def<'a0, 'a1, S0, S1>(&'a1 self, name: S0, default: S1) -> &'a1 str
    where
        S0: Into<&'a0 str>,
        S1: Into<&'a1 str>,
    {
        if let Some(v) = self.attr(name.into()) {
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
    fn format_boolean(&self, buf: &mut String, b: bool) {
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
    fn format_float(&self, buf: &mut String, f: f64) {
        match self.part_type {
            FormatPartType::Number => {
                let dec = self.attr_def("number:decimal-places", "0").parse::<usize>();
                if let Ok(dec) = dec {
                    buf.push_str(&format!("{:.*}", dec, f));
                }
            }
            FormatPartType::Scientific => {
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
    fn format_str(&self, buf: &mut String, s: &str) {
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
    #[allow(clippy::collapsible_if)]
    fn format_datetime(&self, buf: &mut String, d: &NaiveDateTime, h12: bool) {
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
    fn format_time_duration(&self, buf: &mut String, d: &Duration) {
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

/// Creates a new number format.
pub fn create_boolean_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::Boolean);
    v.push_boolean();
    v
}

/// Creates a new number format.
pub fn create_number_format<S: Into<String>>(name: S, decimal: u8, grouping: bool) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::Number);
    v.push_number(decimal, grouping);
    v
}

/// Creates a new number format with a fixed number of decimal places.
pub fn create_number_format_fixed<S: Into<String>>(
    name: S,
    decimal: u8,
    grouping: bool,
) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::Number);
    v.push_number_fix(decimal, grouping);
    v
}

/// Creates a new percantage format.<
pub fn create_percentage_format<S: Into<String>>(name: S, decimal: u8) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::Percentage);
    v.push_number_fix(decimal, false);
    v.push_text("%");
    v
}

/// Creates a new currency format.
pub fn create_currency_prefix<S1, S2, S3, S4>(
    name: S1,
    country: S2,
    language: S3,
    symbol: S4,
) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
    S3: Into<String>,
    S4: Into<String>,
{
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::Currency);
    v.push_currency(country.into(), language.into(), symbol.into());
    v.push_text(" ");
    v.push_number_fix(2, true);
    v
}

/// Creates a new currency format.
pub fn create_currency_suffix<S1, S2, S3, S4>(
    name: S1,
    country: S2,
    language: S3,
    symbol: S4,
) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
    S3: Into<String>,
    S4: Into<String>,
{
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::Currency);
    v.push_number_fix(2, true);
    v.push_text(" ");
    v.push_currency(country.into(), language.into(), symbol.into());
    v
}

/// Creates a new date format D.M.Y
pub fn create_date_dmy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::DateTime);
    v.push_day(FormatNumberStyle::Long);
    v.push_text(".");
    v.push_month(FormatNumberStyle::Long);
    v.push_text(".");
    v.push_year(FormatNumberStyle::Long);
    v
}

/// Creates a new date format M/D/Y
pub fn create_date_mdy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::DateTime);
    v.push_month(FormatNumberStyle::Long);
    v.push_text("/");
    v.push_day(FormatNumberStyle::Long);
    v.push_text("/");
    v.push_year(FormatNumberStyle::Long);
    v
}

/// Creates a datetime format Y-M-D H:M:S
pub fn create_datetime_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::DateTime);
    v.push_day(FormatNumberStyle::Long);
    v.push_text(".");
    v.push_month(FormatNumberStyle::Long);
    v.push_text(".");
    v.push_year(FormatNumberStyle::Long);
    v.push_text(" ");
    v.push_hours(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_minutes(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_seconds(FormatNumberStyle::Long);
    v
}

/// Creates a new time-Duration format H:M:S
pub fn create_time_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_with_name(name.into(), ValueType::TimeDuration);
    v.push_hours(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_minutes(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_seconds(FormatNumberStyle::Long);
    v
}
