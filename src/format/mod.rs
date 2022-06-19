//!
//! Defines ValueFormat for formatting related issues
//!
//! ```
//! use spreadsheet_ods::{ValueFormat, ValueType};
//! use spreadsheet_ods::format::{FormatCalendar, FormatMonth, FormatNumberStyle, FormatTextual};
//!
//! let mut v = ValueFormat::new_named("dt0", ValueType::DateTime);
//! v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
//! v.push_text(".");
//! v.push_month(FormatNumberStyle::Long, FormatTextual::Numeric, FormatMonth::Nominativ, FormatCalendar::Default);
//! v.push_text(".");
//! v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
//! v.push_text(" ");
//! v.push_hours(FormatNumberStyle::Long);
//! v.push_text(":");
//! v.push_minutes(FormatNumberStyle::Long);
//! v.push_text(":");
//! v.push_seconds(FormatNumberStyle::Long, 0);
//!
//! let mut v = ValueFormat::new_named("n3", ValueType::Number);
//! v.push_number(3, false);
//! ```
//! The output formatting is a rough approximation with the possibilities
//! offered by format! and chrono::format. Especially there is no trace of
//! i18n. But on the other hand the formatting rules are applied by LibreOffice
//! when opening the spreadsheet so typically nobody notices this.
//!

mod formatpart;
mod valueformat;

pub use formatpart::*;
pub use valueformat::*;

use crate::ValueType;
use icu_locid::Locale;
use std::fmt::{Display, Formatter};

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

/// Creates a new number format.
pub fn create_loc_boolean_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Boolean);
    v.push_boolean();
    v
}

/// Creates a new number format.
pub fn create_loc_number_format<S: Into<String>>(
    name: S,
    locale: Locale,
    decimal: u8,
    grouping: bool,
) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Number);
    v.push_number(decimal, grouping);
    v
}

/// Creates a new percentage format.
pub fn create_loc_percentage_format<S: Into<String>>(
    name: S,
    locale: Locale,
    decimal: u8,
) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Percentage);
    v.push_number_fix(decimal, false);
    v.push_text("%");
    v
}

/// Creates a new currency format.
pub fn create_loc_currency_prefix<S1, S2>(
    name: S1,
    locale: Locale,
    symbol_locale: Locale,
    symbol: S2,
) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
{
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Currency);
    v.push_currency_symbol(symbol_locale, symbol.into());
    v.push_text(" ");
    v.push_number_fix(2, true);
    v
}

/// Creates a new currency format.
pub fn create_loc_currency_suffix<S1, S2>(
    name: S1,
    locale: Locale,
    symbol_locale: Locale,
    symbol: S2,
) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
{
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Currency);
    v.push_number_fix(2, true);
    v.push_text(" ");
    v.push_currency_symbol(symbol_locale, symbol.into());
    v
}

/// Creates a new date format D.M.Y
pub fn create_loc_date_dmy_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::DateTime);
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text(".");
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text(".");
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v
}

/// Creates a new date format M/D/Y
pub fn create_loc_date_mdy_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::DateTime);
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text("/");
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text("/");
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v
}

/// Creates a datetime format Y-M-D H:M:S
pub fn create_loc_datetime_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::DateTime);
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text(".");
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text(".");
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text(" ");
    v.push_hours(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_minutes(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_seconds(FormatNumberStyle::Long, 0);
    v
}

/// Creates a new time-Duration format H:M:S
pub fn create_loc_time_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::TimeDuration);
    v.push_hours(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_minutes(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_seconds(FormatNumberStyle::Long, 0);
    v
}

/// Creates a new number format.
pub fn create_boolean_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Boolean);
    v.push_boolean();
    v
}

/// Creates a new number format.
pub fn create_number_format<S: Into<String>>(name: S, decimal: u8, grouping: bool) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Number);
    v.push_number(decimal, grouping);
    v
}

/// Creates a new number format with a fixed number of decimal places.
pub fn create_number_format_fixed<S: Into<String>>(
    name: S,
    decimal: u8,
    grouping: bool,
) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Number);
    v.push_number_fix(decimal, grouping);
    v
}

/// Creates a new percentage format.
pub fn create_percentage_format<S: Into<String>>(name: S, decimal: u8) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Percentage);
    v.push_number_fix(decimal, false);
    v.push_text("%");
    v
}

/// Creates a new currency format.
pub fn create_currency_prefix<S1, S2>(name: S1, symbol_locale: Locale, symbol: S2) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
{
    let mut v = ValueFormat::new_named(name.into(), ValueType::Currency);
    v.push_currency_symbol(symbol_locale, symbol.into());
    v.push_text(" ");
    v.push_number_fix(2, true);
    v
}

/// Creates a new currency format.
pub fn create_currency_suffix<S1, S2>(name: S1, symbol_locale: Locale, symbol: S2) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
{
    let mut v = ValueFormat::new_named(name.into(), ValueType::Currency);
    v.push_number_fix(2, true);
    v.push_text(" ");
    v.push_currency_symbol(symbol_locale, symbol.into());
    v
}

/// Creates a new date format YYYY-MM-DD
pub fn create_date_iso_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text("-");
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text("-");
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v
}

/// Creates a new date format D.M.Y
pub fn create_date_dmy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text(".");
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text(".");
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v
}

/// Creates a new date format M/D/Y
pub fn create_date_mdy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text("/");
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text("/");
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v
}

/// Creates a datetime format Y-M-D H:M:S
pub fn create_datetime_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.push_year(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text("-");
    v.push_month(
        FormatNumberStyle::Long,
        FormatTextual::Numeric,
        FormatMonth::Nominativ,
        FormatCalendar::Default,
    );
    v.push_text("-");
    v.push_day(FormatNumberStyle::Long, FormatCalendar::Default);
    v.push_text(" ");
    v.push_hours(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_minutes(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_seconds(FormatNumberStyle::Long, 0);
    v
}

/// Creates a new time-Duration format H:M:S
pub fn create_time_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::TimeDuration);
    v.push_hours(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_minutes(FormatNumberStyle::Long);
    v.push_text(":");
    v.push_seconds(FormatNumberStyle::Long, 0);
    v
}
