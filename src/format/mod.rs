//!
//! Defines ValueFormat for formatting related issues
//!
//! ```
//! use spreadsheet_ods::{ValueFormat, ValueType};
//! use spreadsheet_ods::format::{FormatCalendarStyle, FormatNumberStyle};
//!
//! let mut v = ValueFormat::new_named("dt0", ValueType::DateTime);
//! v.part_day().long_style().build();
//! v.part_text(".");
//! v.part_month().long_style().build();
//! v.part_text(".");
//! v.part_year().long_style().build();
//! v.part_text(" ");
//! v.part_hours().long_style().build();
//! v.part_text(":");
//! v.part_minutes().long_style().build();
//! v.part_text(":");
//! v.part_seconds().long_style().build();
//!
//! let mut v = ValueFormat::new_named("n3", ValueType::Number);
//! v.part_number().decimal_places(3);
//! ```
//!

mod builder;
mod create;

pub use builder::*;
pub use create::*;

use crate::attrmap2::AttrMap2;
use crate::style::stylemap::StyleMap;
use crate::style::units::{
    Angle, FontSize, FontStyle, FontVariant, FontWeight, FormatSource, Length, LetterSpacing,
    LineMode, LineStyle, LineType, LineWidth, Percent, RotationScale, TextCombine, TextCondition,
    TextDisplay, TextEmphasize, TextEmphasizePosition, TextPosition, TextRelief, TextTransform,
    TransliterationStyle,
};
use crate::style::ParseStyleAttr;
use crate::style::{
    color_string, shadow_string, text_position, StyleOrigin, StyleUse, TextStyleRef,
};
use crate::{OdsError, ValueType};
use color::Rgb;
use icu_locid::subtags::{Language, Region, Script};
use icu_locid::{LanguageIdentifier, Locale};
use std::fmt::{Display, Formatter};
use std::str::FromStr;

style_ref!(ValueFormatRef);

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

// TODO: Is there a better way to describe this? All the "is valid for"s are somewhat annoying...

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

impl ValueFormat {
    /// New, empty.
    pub fn new_empty() -> Self {
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
    pub(crate) fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// Text style attributes.
    pub(crate) fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    number_title!(attr);
    number_locale!(attr);
    number_transliteration_locale!(attr);
    number_transliteration_format!(attr);
    number_transliteration_style!(attr);
    style_volatile!(attr);
    number_automatic_order!(attr);
    number_format_source!(attr);
    number_truncate_on_overflow!(attr);

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
    #[must_use]
    pub fn part_fill_character(&mut self) -> PartFillCharacterBuilder<'_> {
        PartFillCharacterBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:scientific-number element specifies the display formatting properties for a
    /// number style that should be displayed in scientific format.
    ///
    /// Can be used with ValueTypes:
    /// * Number
    #[must_use]
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
    #[must_use]
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
    #[must_use]
    pub fn part_currency(&mut self) -> PartCurrencySymbolBuilder<'_> {
        PartCurrencySymbolBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:day element specifies a day of a month in a date.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
    pub fn part_day(&mut self) -> PartDayBuilder<'_> {
        PartDayBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:month element specifies a month in a date.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
    pub fn part_month(&mut self) -> PartMonthBuilder<'_> {
        PartMonthBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:year element specifies a year in a date
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
    pub fn part_year(&mut self) -> PartYearBuilder<'_> {
        PartYearBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:era element specifies an era in which a year is counted
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
    pub fn part_era(&mut self) -> PartEraBuilder<'_> {
        PartEraBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:day-of-week element specifies a day of a week in a date
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
    pub fn part_day_of_week(&mut self) -> PartDayOfWeekBuilder<'_> {
        PartDayOfWeekBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:week-of-year element specifies a week of a year in a date.
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
    pub fn part_week_of_year(&mut self) -> PartWeekOfYearBuilder<'_> {
        PartWeekOfYearBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:quarter element specifies a quarter of the year in a date
    ///
    /// Can be used with ValueTypes:
    /// * DateTime
    #[must_use]
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
    #[must_use]
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
    #[must_use]
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
    #[must_use]
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
    #[must_use]
    pub fn part_am_pm(&mut self) -> PartAmPmBuilder<'_> {
        PartAmPmBuilder::new(self)
    }

    /// Adds a format part to this format.
    ///
    /// The number:boolean element marks the position of the Boolean value of a Boolean style.
    ///
    /// Can be used with ValueTypes:
    /// * Boolean
    #[must_use]
    pub fn part_boolean(&mut self) -> PartBooleanBuilder<'_> {
        PartBooleanBuilder::new(self)
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
    #[must_use]
    pub fn part_text<S: Into<String>>(&mut self, text: S) -> PartTextBuilder<'_> {
        PartTextBuilder::new(self).text(text.into())
    }

    /// Adds a format part to this format.
    ///    
    /// The number:text-content element marks the position of variable text content of a text
    /// style.
    ///
    /// Can be used with ValueTypes:
    /// * Text
    #[must_use]
    pub fn part_text_content(&mut self) -> PartTextContentBuilder<'_> {
        PartTextContentBuilder::new(self)
    }

    /// Use part_number instead.
    #[deprecated]
    pub fn push_number(&mut self, decimal_places: u8, grouping: bool) {
        self.part_number()
            .decimal_places(decimal_places)
            .if_then(grouping, |p| p.grouping())
            .min_decimal_places(0)
            .min_integer_digits(1)
            .build();
    }

    /// Use part_number instead.
    #[deprecated]
    pub fn push_number_fix(&mut self, decimal_places: u8, grouping: bool) {
        self.part_number()
            .fixed_decimal_places(decimal_places)
            .if_then(grouping, |p| p.grouping())
            .min_integer_digits(1)
            .build();
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
            .if_then(grouping, |p| p.grouping())
            .build();
    }

    /// Use part_scientific instead.
    #[deprecated]
    pub fn push_scientific(&mut self, decimal_places: u8) {
        self.part_scientific()
            .decimal_places(decimal_places)
            .build();
    }

    /// Use part_currency instead.
    #[deprecated]
    pub fn push_currency_symbol<S>(&mut self, locale: Locale, symbol: S)
    where
        S: Into<String>,
    {
        self.part_currency().locale(locale).symbol(symbol).build();
    }

    /// Use part_day instead.
    #[deprecated]
    pub fn push_day(&mut self, style: FormatNumberStyle) {
        self.part_day().style(style).build();
    }

    /// Use part_month instead.
    #[deprecated]
    pub fn push_month(&mut self, style: FormatNumberStyle, textual: bool) {
        self.part_month()
            .style(style)
            .if_then(textual, |p| p.textual())
            .build();
    }

    /// Use part_year instead.
    #[deprecated]
    pub fn push_year(&mut self, style: FormatNumberStyle) {
        self.part_year().style(style).build();
    }

    /// Use part_era instead.
    #[deprecated]
    pub fn push_era(&mut self, style: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.part_era().style(style).calendar(calendar).build();
    }

    /// Use part_day_of_week instead.
    #[deprecated]
    pub fn push_day_of_week(&mut self, style: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.part_day_of_week()
            .style(style)
            .calendar(calendar)
            .build();
    }

    /// Use part_week_of_year instead.
    #[deprecated]
    pub fn push_week_of_year(&mut self, calendar: FormatCalendarStyle) {
        self.part_week_of_year().calendar(calendar).build();
    }

    /// Use part_quarter instead.
    #[deprecated]
    pub fn push_quarter(&mut self, style: FormatNumberStyle, calendar: FormatCalendarStyle) {
        self.part_quarter().style(style).calendar(calendar).build();
    }

    /// Use part_hours instead.
    #[deprecated]
    pub fn push_hours(&mut self, style: FormatNumberStyle) {
        self.part_hours().style(style).build();
    }

    /// Use part_minutes instead.
    #[deprecated]
    pub fn push_minutes(&mut self, style: FormatNumberStyle) {
        self.part_minutes().style(style).build();
    }

    /// Use part_seconds.
    #[deprecated]
    pub fn push_seconds(&mut self, style: FormatNumberStyle, decimal_places: u8) {
        self.part_seconds()
            .style(style)
            .decimal_places(decimal_places)
            .build();
    }

    /// Use part_am_pm instead.
    #[deprecated]
    pub fn push_am_pm(&mut self) {
        self.part_am_pm().build();
    }

    /// Use part_boolean instead.
    #[deprecated]
    pub fn push_boolean(&mut self) {
        self.part_boolean().build();
    }

    /// Use part_text instead.
    #[deprecated]
    pub fn push_text<S: Into<String>>(&mut self, text: S) {
        self.part_text(text).build();
    }

    /// Use part_text_content instead.
    #[deprecated]
    pub fn push_text_content(&mut self) {
        self.part_text_content().build();
    }

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
    /// Textposition for embedded text when acting as a number format part.
    ///
    /// The number:position attribute specifies the position where text appears.
    /// The index of a position starts with 1 and is counted by digits from right to left in the integer part of
    /// a number, starting left from a decimal separator if one exists, or from the last digit of the number.
    /// Text is inserted before the digit at the specified position. If the value of number:position
    /// attribute is greater than the value of number:min-integer-digits 19.355 and greater than
    /// the number of integer digits in the number, text is prepended to the number.
    position: u32,
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
            position: 0,
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
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
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

    /// Sets the position for embedded text in a number format part.
    pub fn set_position(&mut self, pos: u32) {
        self.position = pos;
    }

    /// The position for embedded text in a number format part.
    pub fn position(&self) -> u32 {
        self.position
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
