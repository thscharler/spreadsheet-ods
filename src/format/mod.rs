//!
//! Defines ValueFormat for formatting related issues
//!
//! ```
//! use spreadsheet_ods::{ValueFormat, ValueType};
//! use spreadsheet_ods::format::{FormatCalendarStyle, FormatMonth, FormatNumberStyle, FormatTextual};
//!
//! let mut v = ValueFormat::new_named("dt0", ValueType::DateTime);
//! v.part_day().long_style().push();
//! v.part_text(".");
//! v.part_month().long_style().push();
//! v.part_text(".");
//! v.part_year().long_style().push();
//! v.part_text(" ");
//! v.part_hours().long_style().push();
//! v.part_text(":");
//! v.part_minutes().long_style().push();
//! v.part_text(":");
//! v.part_seconds().long_style().push();
//!
//! let mut v = ValueFormat::new_named("n3", ValueType::Number);
//! v.part_number().decimal_places(3);
//! ```
//!
//! X!!
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
use color::Rgb;
use icu_locid::subtags::{Language, Region, Script};
use icu_locid::{LanguageIdentifier, Locale};
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

/// Format source
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FormatSource {
    Fixed,
    Language,
}

impl Display for FormatSource {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            FormatSource::Fixed => write!(f, "fixed"),
            FormatSource::Language => write!(f, "language"),
        }
    }
}

impl FromStr for FormatSource {
    type Err = OdsError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "fixed" => Ok(FormatSource::Fixed),
            "language" => Ok(FormatSource::Language),
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
// ok number:automatic-order 19.340
//
// ValueType:Percentage -> number:percentage-style
//      no extras
//
// ValueType:DateTime -> number:date-style
// ok number:automaticorder 19.340
// ok number:format-source 19.347,
//
// ValueType:TimeDuration -> number:time-style
// ok number:format-source 19.347
// ok number:truncate-on-overflow 19.365
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
    #[deprecated]
    pub fn new_with_name<S: Into<String>>(name: S, value_type: ValueType) -> Self {
        Self::new_named(name, value_type)
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
        v.set_locale(locale);
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

    /// The style:name attribute specifies names that reference style mechanisms.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// The number:title attribute specifies the title of a data style.
    pub fn set_title<S: Into<String>>(&mut self, title: S) {
        self.attr.set_attr("number:title", title.into());
    }

    /// The number:title attribute specifies the title of a data style.
    pub fn title(&self) -> Option<&String> {
        self.attr.attr("number:title")
    }

    /// The style:display-name attribute specifies the name of a style as it should appear in the user
    /// interface. If this attribute is not present, the display name should be the same as the style name.
    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("number:country", name.into());
    }

    /// The style:display-name attribute specifies the name of a style as it should appear in the user
    /// interface. If this attribute is not present, the display name should be the same as the style name.
    pub fn display_name(&self) -> Option<&String> {
        self.attr.attr("number:country")
    }

    /// The number:language attribute specifies a language code. The country code is used for
    /// formatting properties whose evaluation is locale-dependent.
    /// If a language code is not specified, either the system settings or the setting for the system's
    /// language are used, depending on the property whose value should be evaluated.
    ///
    /// The number:country attribute specifies a country code for a data style. The country code is
    /// used for formatting properties whose evaluation is locale-dependent.
    /// If a country is not specified, the system settings are used.
    ///
    /// The number:script attribute specifies a script code. The script code is used for formatting
    /// properties whose evaluation is locale-dependent. The attribute should be used only if necessary
    /// according to the rules of §2.2.3 of [RFC5646](https://datatracker.ietf.org/doc/html/rfc5646), or its successors.
    pub fn set_locale(&mut self, locale: Locale) {
        if locale != Locale::UND {
            self.attr
                .set_attr("number:language", locale.id.language.to_string());
            if let Some(region) = locale.id.region {
                self.attr.set_attr("number:country", region.to_string());
            } else {
                self.attr.clear_attr("number:country");
            }
            if let Some(script) = locale.id.script {
                self.attr.set_attr("number:script", script.to_string());
            } else {
                self.attr.clear_attr("number:script");
            }
        } else {
            self.attr.clear_attr("number:language");
            self.attr.clear_attr("number:country");
            self.attr.clear_attr("number:script");
        }
    }

    /// Returns number:language, number:country and number:script as a locale.
    pub fn locale(&self) -> Option<Locale> {
        if let Some(language) = self.attr.attr("number:language") {
            if let Some(language) = Language::from_bytes(language.as_bytes()).ok() {
                let region = if let Some(region) = self.attr.attr("number:country") {
                    Region::from_bytes(region.as_bytes()).ok()
                } else {
                    None
                };
                let script = if let Some(script) = self.attr.attr("number:script") {
                    Script::from_bytes(script.as_bytes()).ok()
                } else {
                    None
                };

                let id = LanguageIdentifier::from((language, script, region));

                Some(Locale::from(id))
            } else {
                None
            }
        } else {
            None
        }
    }

    /// The number:transliteration-language attribute specifies a language code in
    /// conformance with [RFC5646](https://datatracker.ietf.org/doc/html/rfc5646).
    /// If no language/country (locale) combination is specified, the locale of the data style is used
    ///
    /// The number:transliteration-country attribute specifies a country code in conformance
    /// with [RFC5646](https://datatracker.ietf.org/doc/html/rfc5646).
    /// If no language/country (locale) combination is specified, the locale of the data style is used.
    pub fn set_transliteration_locale(&mut self, locale: Locale) {
        if locale != Locale::UND {
            self.attr.set_attr(
                "number:transliteration-language",
                locale.id.language.to_string(),
            );
            if let Some(region) = locale.id.region {
                self.attr
                    .set_attr("number:transliteration-country", region.to_string());
            } else {
                self.attr.clear_attr("number:transliteration-country");
            }
        } else {
            self.attr.clear_attr("number:transliteration-language");
            self.attr.clear_attr("number:transliteration-country");
        }
    }

    /// Returns number:transliteration_language and number:transliteration_country as a locale.
    pub fn transliteration_locale(&self) -> Option<Locale> {
        if let Some(language) = self.attr.attr("number:language") {
            if let Some(language) = Language::from_bytes(language.as_bytes()).ok() {
                let region = if let Some(region) = self.attr.attr("number:country") {
                    Region::from_bytes(region.as_bytes()).ok()
                } else {
                    None
                };

                let id = LanguageIdentifier::from((language, None, region));

                Some(Locale::from(id))
            } else {
                None
            }
        } else {
            None
        }
    }

    /// The number:transliteration-format attribute specifies which number characters to use.
    ///
    /// The value of the number:transliteration-format attribute shall be a decimal "DIGIT ONE"
    /// character with numeric value 1 as listed in the Unicode Character Database file UnicodeData.txt
    /// with value 'Nd' (Numeric decimal digit) in the General_Category/Numeric_Type property field 6
    /// and value '1' in the Numeric_Value fields 7 and 8, respectively as listed in
    /// DerivedNumericValues.txt
    ///
    /// If no format is specified the default ASCII representation of Latin-Indic digits is used, other
    /// transliteration attributes present in that case are ignored.
    ///
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
    ///
    /// The semantics of the values of the number:transliteration-style attribute are locale- and
    /// implementation-dependent.
    ///
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
    /// * false: consumers should discard the unused styles.
    /// * true: consumers should keep unused styles.
    pub fn set_volatile(&mut self, volatile: bool) {
        self.attr.set_attr("style:volatile", volatile.to_string());
    }

    /// Volatile format.
    pub fn volatile(&self) -> Option<bool> {
        match self.attr.attr("style:volatile") {
            None => None,
            Some(s) => FromStr::from_str(s.as_str()).ok(),
        }
    }

    /// The number:automatic-order attribute specifies whether data is ordered to match the default
    /// order for the language and country of a data style.
    /// The defined values for the number:automatic-order attribute are:
    /// * false: data is not ordered to match the default order for the language and country of a data
    /// style.
    /// * true: data is ordered to match the default order for the language and country of a data style.
    /// The default value for this attribute is false.
    ///
    /// This attribute is valid for ValueType::DateTime and ValueType::TimeDuration.
    pub fn set_automatic_order(&mut self, volatile: bool) {
        self.attr
            .set_attr("number:automatic-order", volatile.to_string());
    }

    /// Automatic order.
    pub fn automatic_order(&self) -> Option<bool> {
        if let Some(v) = self.attr.attr("number:automatic-order") {
            v.parse().ok()
        } else {
            None
        }
    }

    /// The number:format-source attribute specifies the source of definitions of the short and
    /// long display formats.
    ///
    /// The defined values for the number:format-source attribute are:
    /// * fixed: the values short and long of the number:style attribute are defined by this
    /// standard.
    /// * language: the meaning of the values long and short of the number:style attribute
    /// depend upon the number:language and number:country attributes of the date style. If
    /// neither of those attributes are specified, consumers should use their default locale for short
    /// and long date and time formats.
    ///
    /// The default value for this attribute is fixed.
    ///
    /// This attribute is valid for ValueType::DateTime and ValueType::TimeDuration.
    pub fn set_format_source(&mut self, source: FormatSource) {
        self.attr
            .set_attr("number:format-source", source.to_string());
    }

    /// The source of definitions of the short and long display formats.
    pub fn format_source(&mut self) -> Option<FormatSource> {
        if let Some(v) = self.attr.attr("number:format-source") {
            v.parse().ok()
        } else {
            None
        }
    }

    /// The number:truncate-on-overflow attribute specifies if a time or duration for which the
    /// value to be displayed by the largest time component specified in the style is too large to be
    /// displayed using the value range for number:hours 16.29.20 (0 to 23), or
    /// number:minutes 16.29.21 or number:seconds 16.29.22 (0 to 59) is truncated or if the
    /// value range of this component is extended. The largest time component is those for which a value
    /// of "1" represents the longest period of time.
    /// If a value gets truncated, then its value is displayed modulo 24 (for number:hours) or modulo
    /// 60 (for number:minutes and number:seconds).
    ///
    /// If the value range of a component get extended, then values larger than 23 or 59 are displayed.
    /// The defined values for the number:truncate-on-overflow element are:
    /// * false: the value range of the component is extended.
    /// * true: the value range of the component is not extended.
    ///
    /// The default value for this attribute is true.
    ///
    /// This attribute is valid for ValueType::TimeDuration.
    pub fn set_truncate_on_overflow(&mut self, truncate: bool) {
        self.attr
            .set_attr("number:truncate-on-overflow", truncate.to_string());
    }

    /// Truncate time-values on overflow.
    pub fn truncate_on_overflow(&mut self) -> Option<bool> {
        if let Some(v) = self.attr.attr("number:truncate-on-overflow") {
            v.parse().ok()
        } else {
            None
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

    /// Sets the storage location for this ValueFormat. Either content.xml
    /// or styles.xml.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Returns the storage location.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// How is the style used in the document.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// How is the style used in the document.
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
    pub fn push_era(&mut self, style: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.part_era().style(style).calendar(calendar).push();
    }

    /// Use part_day_of_week instead.
    #[deprecated]
    pub fn push_day_of_week(&mut self, style: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.part_day_of_week()
            .style(style)
            .calendar(calendar)
            .push();
    }

    /// Use part_week_of_year instead.
    #[deprecated]
    pub fn push_week_of_year(&mut self, calendar: FormatCalendarStyle) {
        self.part_week_of_year().calendar(calendar).push();
    }

    /// Use part_quarter instead.
    #[deprecated]
    pub fn push_quarter(&mut self, style: FormatNumberStyle, calendar: FormatCalendarStyle) {
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
pub enum FormatCalendarStyle {
    Gregorian,
    Gengou,
    Roc,
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
            FormatCalendarStyle::Roc => write!(f, "ROC"),
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
    pub fn attr_def<'a, 'b, S>(&'a self, name: &'b str, default: S) -> &'a str
    where
        S: Into<&'a str>,
    {
        self.attr.attr_def(name, default)
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
}
