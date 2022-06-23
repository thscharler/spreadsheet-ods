//!
//! Defines ValueFormat for formatting related issues
//!
//! ```
//! use spreadsheet_ods::{ValueFormat, ValueType};
//! use spreadsheet_ods::format::{FormatCalendar, FormatMonth, FormatNumberStyle, FormatTextual};
//!
//! let mut v = ValueFormat::new_named("dt0", ValueType::DateTime);
//! v.part_day().long().push();
//! v.part_text(".");
//! v.part_month().long().push();
//! v.part_text(".");
//! v.part_year().long().push();
//! v.part_text(" ");
//! v.part_hours().long().push();
//! v.part_text(":");
//! v.part_minutes().long().push();
//! v.part_text(":");
//! v.part_seconds().long().push();
//!
//! let mut v = ValueFormat::new_named("n3", ValueType::Number);
//! v.part_number().decimal_places(3);
//! ```
//!
//! The output formatting is a rough approximation with the possibilities
//! offered by format! and chrono::format. Especially there is no trace of
//! i18n. But on the other hand the formatting rules are applied by LibreOffice
//! when opening the spreadsheet so typically nobody notices this.
//!

mod builder;
mod create;

pub use builder::*;
pub use create::*;

use crate::attrmap2::AttrMap2;
use crate::format::{
    PartCurrencySymbolBuilder, PartDayBuilder, PartDayOfWeekBuilder, PartEraBuilder,
    PartFractionBuilder, PartHoursBuilder, PartMinutesBuilder, PartMonthBuilder, PartNumberBuilder,
    PartQuarterBuilder, PartScientificBuilder, PartSecondsBuilder, PartWeekOfYearBuilder,
    PartYearBuilder,
};
use crate::style::stylemap::StyleMap;
use crate::style::units::{
    FontStyle, FontWeight, Length, LineMode, LineStyle, LineType, LineWidth, TextPosition,
    TextRelief, TextTransform,
};
use crate::style::{
    color_string, percent_string, shadow_string, StyleOrigin, StyleUse, TextStyleRef,
};
use crate::{OdsError, ValueType};
use chrono::{Duration, NaiveDateTime};
use color::Rgb;
use icu_locid::subtags::{Language, Region, Script};
use icu_locid::Locale;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

/// Error type for any formatting errors.
#[derive(Debug)]
#[allow(missing_docs)]
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

style_ref!(ValueFormatRef);

/// Transliteration style
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum TransliterationStyle {
    Short,
    Medium,
    Long,
}

impl Display for TransliterationStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            TransliterationStyle::Short => write!(f, "short"),
            TransliterationStyle::Medium => write!(f, "medium"),
            TransliterationStyle::Long => write!(f, "long"),
        }
    }
}

impl FromStr for TransliterationStyle {
    type Err = OdsError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "short" => Ok(TransliterationStyle::Short),
            "medium" => Ok(TransliterationStyle::Medium),
            "long" => Ok(TransliterationStyle::Long),
            _ => Err(OdsError::Parse(s.to_string())),
        }
    }
}

// Styles and attributes
//
// Attributes for all styles:
// ok number:country 19.342
// ok number:language 19.351
// ignore number:rfc-language-tag 19.360
// ok number:script 19.361
// ok number:title 19.364
// ok number:transliteration-country 19.365
// ok number:transliteration-format 19.366
// ok number:transliteration-language 19.367
// ok number:transliteration-style 19.368
// ok style:display-name 19.476
// ok style:name 19.502
// ok style:volatile 19.521
//
// ValueType:Number -> number:number-style
//      no extras
//
// ValueType:Currency -> number:currency-style
// number:automatic-order 19.340
//
// ValueType:Percentage -> number:percentage-style
//      no extras
//
// ValueType:DateTime -> number:date-style
// number:automaticorder 19.340
// number:format-source 19.347,
//
// ValueType:TimeDuration -> number:time-style
// number:format-source 19.347
// number:truncate-on-overflow 19.365
//
// ValueType:Boolean -> number:boolean-style
//      no extras
//
// ValueType:Text -> number:text-style
//      no extras

/// Actual textual formatting of values.
#[derive(Debug, Clone)]
pub struct ValueFormat {
    /// Name
    name: String,
    /// Value type
    v_type: ValueType,
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

impl Default for ValueFormat {
    fn default() -> Self {
        ValueFormat::new()
    }
}

impl ValueFormat {
    /// New, empty.
    pub fn new() -> Self {
        ValueFormat {
            name: String::from(""),
            v_type: ValueType::Text,
            origin: StyleOrigin::Styles,
            styleuse: StyleUse::Default,
            attr: Default::default(),
            textstyle: Default::default(),
            parts: Default::default(),
            stylemaps: None,
        }
    }

    /// New, with name.
    pub fn new_named<S: Into<String>>(name: S, value_type: ValueType) -> Self {
        assert_ne!(value_type, ValueType::Empty);
        ValueFormat {
            name: name.into(),
            v_type: value_type,
            origin: StyleOrigin::Styles,
            styleuse: StyleUse::Default,
            attr: Default::default(),
            textstyle: Default::default(),
            parts: Default::default(),
            stylemaps: None,
        }
    }

    /// New, with name.
    pub fn new_localized<S: Into<String>>(name: S, locale: Locale, value_type: ValueType) -> Self {
        assert_ne!(value_type, ValueType::Empty);
        let mut v = ValueFormat {
            name: name.into(),
            v_type: value_type,
            origin: StyleOrigin::Styles,
            styleuse: StyleUse::Default,
            attr: Default::default(),
            textstyle: Default::default(),
            parts: Default::default(),
            stylemaps: None,
        };
        v.set_language(locale.id.language);
        if let Some(region) = locale.id.region {
            v.set_country(region);
        }
        if let Some(script) = locale.id.script {
            v.set_script(script);
        }
        v
    }

    /// Returns a reference name for this value format.
    pub fn format_ref(&self) -> ValueFormatRef {
        ValueFormatRef::from(self.name().as_str())
    }

    /// The style:name attribute specifies names that reference style mechanisms.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Style name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// The number:title attribute specifies the title of a data style.
    pub fn set_title<S: Into<String>>(&mut self, title: S) {
        self.attr.set_attr("number:title", title.into());
    }

    /// Title
    pub fn title(&self) -> Option<&String> {
        self.attr.attr("number:title")
    }

    /// The style:display-name attribute specifies the name of a style as it should appear in the user
    /// interface. If this attribute is not present, the display name should be the same as the style name.
    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("number:country", name.into());
    }

    /// Display name.
    pub fn display_name(&self) -> Option<&String> {
        self.attr.attr("number:country")
    }

    /// The number:country attribute specifies a country code for a data style. The country code is
    /// used for formatting properties whose evaluation is locale-dependent.
    /// If a country is not specified, the system settings are used
    pub fn set_country(&mut self, country: Region) {
        self.attr.set_attr("number:country", country.to_string());
    }

    /// Country
    pub fn country(&self) -> Option<Region> {
        match self.attr.attr("number:country") {
            None => None,
            Some(v) => v.parse().ok(),
        }
    }

    /// The number:language attribute specifies a language code. The country code is used for
    /// formatting properties whose evaluation is locale-dependent.
    /// If a language code is not specified, either the system settings or the setting for the system's
    /// language are used, depending on the property whose value should be evaluated.
    pub fn set_language(&mut self, language: Language) {
        self.attr.set_attr("number:language", language.to_string());
    }

    /// Language
    pub fn language(&self) -> Option<Language> {
        match self.attr.attr("number:language") {
            None => None,
            Some(v) => v.parse().ok(),
        }
    }

    /// The number:script attribute specifies a script code. The script code is used for formatting
    /// properties whose evaluation is locale-dependent. The attribute should be used only if necessary
    /// according to the rules of §2.2.3 of [RFC5646](https://datatracker.ietf.org/doc/html/rfc5646), or its successors.
    pub fn set_script(&mut self, script: Script) {
        self.attr.set_attr("number:script", script.to_string());
    }

    /// Script
    pub fn script(&self) -> Option<Script> {
        match self.attr.attr("number:script") {
            None => None,
            Some(v) => v.parse().ok(),
        }
    }

    /// The number:transliteration-country attribute specifies a country code in conformance
    /// with [RFC5646](https://datatracker.ietf.org/doc/html/rfc5646).
    /// If no language/country (locale) combination is specified, the locale of the data style is used.
    pub fn set_transliteration_country(&mut self, country: Region) {
        self.attr
            .set_attr("number:transliteration-country", country.to_string());
    }

    /// Transliteration country.
    pub fn transliteration_country(&self) -> Option<Region> {
        match self.attr.attr("number:transliteration-country") {
            None => None,
            Some(v) => v.parse().ok(),
        }
    }

    /// The number:transliteration-language attribute specifies a language code in
    /// conformance with [RFC5646](https://datatracker.ietf.org/doc/html/rfc5646).
    /// If no language/country (locale) combination is specified, the locale of the data style is used
    pub fn set_transliteration_language(&mut self, language: Language) {
        self.attr
            .set_attr("number:transliteration-language", language.to_string());
    }

    /// Transliteration language.
    pub fn transliteration_language(&self) -> Option<Language> {
        match self.attr.attr("number:transliteration-language") {
            None => None,
            Some(v) => v.parse().ok(),
        }
    }

    /// The number:transliteration-format attribute specifies which number characters to use.
    /// The value of the number:transliteration-format attribute shall be a decimal "DIGIT ONE"
    /// character with numeric value 1 as listed in the Unicode Character Database file UnicodeData.txt
    /// with value 'Nd' (Numeric decimal digit) in the General_Category/Numeric_Type property field 6
    /// and value '1' in the Numeric_Value fields 7 and 8, respectively as listed in
    /// DerivedNumericValues.txt
    /// If no format is specified the default ASCII representation of Latin-Indic digits is used, other
    /// transliteration attributes present in that case are ignored.
    /// The default value for this attribute is 1
    pub fn set_transliteration_format(&mut self, format: char) {
        self.attr
            .set_attr("number:transliteration-format", format.into());
    }

    /// Transliteration format.
    pub fn transliteration_format(&self) -> Option<char> {
        match self.attr.attr("number:transliteration-format") {
            None => None,
            Some(v) => v.chars().next(),
        }
    }

    /// The number:transliteration-style attribute specifies the transliteration format of a
    /// number system.
    /// The semantics of the values of the number:transliteration-style attribute are locale- and
    /// implementation-dependent.
    /// The default value for this attribute is short.
    pub fn set_transliteration_style(&mut self, style: TransliterationStyle) {
        self.attr
            .set_attr("number:transliteration-style", style.to_string());
    }

    /// Transliteration style.
    pub fn transliteration_style(&self) -> Option<TransliterationStyle> {
        match self.attr.attr("number:transliteration-style") {
            None => None,
            Some(s) => FromStr::from_str(s.as_str()).ok(),
        }
    }

    /// The style:volatile attribute specifies whether unused style in a document are retained or
    /// discarded by consumers.
    /// The defined values for the style:volatile attribute are:
    ///   false: consumers should discard the unused styles.
    ///   true: consumers should keep unused styles.
    pub fn set_volatile(&mut self, volatile: bool) {
        self.attr.set_attr("style:volatile", volatile.to_string());
    }

    /// Transliteration style.
    pub fn volatile(&self) -> Option<bool> {
        match self.attr.attr("style:volatile") {
            None => None,
            Some(s) => FromStr::from_str(s.as_str()).ok(),
        }
    }

    /// The number:automatic-order attribute specifies whether data is ordered to match the default
    /// order for the language and country of a data style.
    /// The defined values for the number:automatic-order attribute are:
    /// - false: data is not ordered to match the default order for the language and country of a data
    /// style.
    /// - true: data is ordered to match the default order for the language and country of a data style.
    /// The default value for this attribute is false.
    ///
    /// This attribute is valid for date and currency formats.
    pub fn set_automatic_order(&mut self, volatile: bool) {
        self.attr
            .set_attr("number:automatic-order", volatile.to_string());
    }

    /// Automatic order.
    pub fn automatic_order(&self) -> Option<bool> {
        match self.attr.attr("number:automatic-order") {
            None => None,
            Some(s) => FromStr::from_str(s.as_str()).ok(),
        }
    }

    /// Sets the value type.
    pub fn set_value_type(&mut self, value_type: ValueType) {
        assert_ne!(value_type, ValueType::Empty);
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

    /// All direct attributes of the number:xxx-style tag.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// All direct attributes of the number:xxx-style tag.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Text style attributes.
    pub fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// Text style attributes.
    pub fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    text!(textstyle_mut);

    /// Adds a format part to this format.
    ///
    /// The number:number element specifies the display formatting properties for a decimal
    /// number.
    ///
    /// Can be used with ValueTypes:
    /// * Currency
    /// * Number
    /// * Percentage
    pub fn part_number(&mut self) -> PartNumberBuilder<'_> {
        PartNumberBuilder::new(self)
    }

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
    pub fn part_fill_character(&mut self, c: char) {
        let mut part = FormatPart::new(FormatPartType::FillCharacter);
        part.set_content(c.to_string());
        self.push_part(part);
    }

    /// Adds a format part to this format.
    ///
    /// The number:scientific-number element specifies the display formatting properties for a
    /// number style that should be displayed in scientific format.
    ///
    /// Can be used with ValueTypes:
    /// * Number
    pub fn part_scientific(&mut self) -> PartScientificBuilder<'_> {
        PartScientificBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:fraction element specifies the display formatting properties for a number style
    /// that should be displayed as a fraction.
    ///
    /// Can be used with ValueTypes:
    /// * Number
    pub fn part_fraction(&mut self) -> PartFractionBuilder<'_> {
        PartFractionBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:currency-symbol element specifies whether a currency symbol is displayed in
    /// a currency style.
    ///
    /// Can be used with ValueTypes:
    /// * Currency
    pub fn part_currency(&mut self) -> PartCurrencySymbolBuilder<'_> {
        PartCurrencySymbolBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:day element specifies a day of a month in a date.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_day(&mut self) -> PartDayBuilder<'_> {
        PartDayBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:month element specifies a month in a date.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_month(&mut self) -> PartMonthBuilder<'_> {
        PartMonthBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:year element specifies a year in a date
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_year(&mut self) -> PartYearBuilder<'_> {
        PartYearBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:era element specifies an era in which a year is counted
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_era(&mut self) -> PartEraBuilder<'_> {
        PartEraBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:day-of-week element specifies a day of a week in a date
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_day_of_week(&mut self) -> PartDayOfWeekBuilder<'_> {
        PartDayOfWeekBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:week-of-year element specifies a week of a year in a date.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_week_of_year(&mut self) -> PartWeekOfYearBuilder<'_> {
        PartWeekOfYearBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:quarter element specifies a quarter of the year in a date
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    pub fn part_quarter(&mut self) -> PartQuarterBuilder<'_> {
        PartQuarterBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:hours element specifies whether hours are displayed as part of a date or time.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    /// * TimeDuration
    pub fn part_hours(&mut self) -> PartHoursBuilder<'_> {
        PartHoursBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:minutes element specifies whether minutes are displayed as part of a date or
    /// time.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    /// * TimeDuration
    pub fn part_minutes(&mut self) -> PartMinutesBuilder<'_> {
        PartMinutesBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:seconds element specifies whether seconds are displayed as part of a date or
    /// time.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    /// * TimeDuration
    pub fn part_seconds(&mut self) -> PartSecondsBuilder<'_> {
        PartSecondsBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:am-pm element specifies whether AM/PM is included as part of a date or time.
    /// If a number:am-pm element is contained in a date or time style, hours are displayed using
    /// values from 1 to 12 only.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    /// * TimeDuration
    pub fn part_am_pm(&mut self) {
        let part = FormatPart::new(FormatPartType::AmPm);
        self.push_part(part);
    }

    /// Adds a format part to this format.
    ///
    /// The number:boolean element marks the position of the Boolean value of a Boolean style.
    ///
    /// Can be used with ValueTypes:
    /// * Boolean
    pub fn part_boolean(&mut self) {
        let part = FormatPart::new(FormatPartType::Boolean);
        self.push_part(part);
    }

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
    pub fn part_text<S: Into<String>>(&mut self, txt: S) {
        let mut part = FormatPart::new(FormatPartType::Text);
        part.set_content(txt.into());
        self.push_part(part);
    }

    /// Adds a format part to this format.
    ///    
    /// The number:text-content element marks the position of variable text content of a text
    /// style.
    ///
    /// Can be used with ValueTypes:
    /// * Text
    pub fn part_text_content(&mut self) {
        let part = FormatPart::new(FormatPartType::TextContent);
        self.push_part(part);
    }

    /// Use part_number instead.
    #[deprecated]
    pub fn push_number(&mut self, decimal_places: u8, grouping: bool) {
        self.part_number()
            .decimal_places(decimal_places)
            .test(grouping, |p| p.grouping())
            .min_decimal_places(0)
            .min_integer_digits(1)
            .push();
    }

    /// Use part_number instead.
    #[deprecated]
    pub fn push_number_fix(&mut self, decimal_places: u8, grouping: bool) {
        self.part_number()
            .fixed_decimal_places(decimal_places)
            .test(grouping, |p| p.grouping())
            .min_integer_digits(1)
            .push();
    }

    /// Use part_fraction instead.
    #[deprecated]
    pub fn push_fraction(
        &mut self,
        denominator: i64,
        min_denominator_digits: u8,
        min_integer_digits: u8,
        min_numerator_digits: u8,
        grouping: bool,
    ) {
        self.part_fraction()
            .denominator(denominator)
            .min_denominator_digits(min_denominator_digits)
            .min_integer_digits(min_integer_digits)
            .min_numerator_digits(min_numerator_digits)
            .test(grouping, |p| p.grouping())
            .push();
    }

    /// Use part_scientific instead.
    #[deprecated]
    pub fn push_scientific(&mut self, decimal_places: u8) {
        self.part_scientific().decimal_places(decimal_places).push();
    }

    /// Use part_currency instead.
    #[deprecated]
    pub fn push_currency_symbol<S>(&mut self, locale: Locale, symbol: S)
    where
        S: Into<String>,
    {
        self.part_currency().locale(locale).symbol(symbol).push();
    }

    /// Use part_day instead.
    #[deprecated]
    pub fn push_day(&mut self, style: FormatNumberStyle) {
        self.part_day().style(style).push();
    }

    /// Use part_month instead.
    #[deprecated]
    pub fn push_month(&mut self, style: FormatNumberStyle, textual: bool) {
        self.part_month()
            .style(style)
            .test(textual, |p| p.textual())
            .push();
    }

    /// Use part_year instead.
    #[deprecated]
    pub fn push_year(&mut self, style: FormatNumberStyle) {
        self.part_year().style(style).push();
    }

    /// Use part_era instead.
    #[deprecated]
    pub fn push_era(&mut self, style: FormatNumberStyle, calendar: FormatCalendar) {
        self.part_era().style(style).calendar(calendar).push();
    }

    /// Use part_day_of_week instead.
    #[deprecated]
    pub fn push_day_of_week(&mut self, style: FormatNumberStyle, calendar: FormatCalendar) {
        self.part_day_of_week()
            .style(style)
            .calendar(calendar)
            .push();
    }

    /// Use part_week_of_year instead.
    #[deprecated]
    pub fn push_week_of_year(&mut self, calendar: FormatCalendar) {
        self.part_week_of_year().calendar(calendar).push();
    }

    /// Use part_quarter instead.
    #[deprecated]
    pub fn push_quarter(&mut self, style: FormatNumberStyle, calendar: FormatCalendar) {
        self.part_quarter().style(style).calendar(calendar).push();
    }

    /// Use part_hours instead.
    #[deprecated]
    pub fn push_hours(&mut self, style: FormatNumberStyle) {
        self.part_hours().style(style).push();
    }

    /// Use part_minutes instead.
    #[deprecated]
    pub fn push_minutes(&mut self, style: FormatNumberStyle) {
        self.part_minutes().style(style).push();
    }

    /// Use part_seconds.
    #[deprecated]
    pub fn push_seconds(&mut self, style: FormatNumberStyle, decimal_places: u8) {
        self.part_seconds()
            .style(style)
            .decimal_places(decimal_places)
            .push();
    }

    /// Use part_am_pm instead.
    #[deprecated]
    pub fn push_am_pm(&mut self) {
        self.part_am_pm();
    }

    /// Use part_boolean instead.
    #[deprecated]
    pub fn push_boolean(&mut self) {
        self.part_boolean();
    }

    /// Use part_text instead.
    #[deprecated]
    pub fn push_text<S: Into<String>>(&mut self, text: S) {
        self.part_text(text);
    }

    /// Use part_text_content instead.
    #[deprecated]
    pub fn push_text_content(&mut self) {
        self.part_text_content();
    }

    // /// The number:///-text element specifies text that is displayed at one specific position
    // /// within a number.
    // pub fn push_embedded_text<S: Into<String>>(&mut self, position: u8, text: S) {
    //     self.push_part(FormatPart::new_embedded_text(position, text));
    // }

    /// Adds a format part.
    pub fn push_part(&mut self, part: FormatPart) {
        self.parts.push(part);
    }

    /// Adds all format parts.
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
            .any(|v| v.part_type() == FormatPartType::AmPm);

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

/// Calendar types.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatCalendar {
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
