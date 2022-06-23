use crate::format::FormatNumberStyle;
use crate::{ValueFormat, ValueType};
use icu_locid::Locale;

/// Creates a new number format.
pub fn create_loc_boolean_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Boolean);
    v.part_boolean();
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
    v.part_number()
        .decimal_places(decimal)
        .test(grouping, |p| p.grouping())
        .push();
    // v.push_number(decimal, grouping);
    v
}

/// Creates a new percentage format.
pub fn create_loc_percentage_format<S: Into<String>>(
    name: S,
    locale: Locale,
    decimal: u8,
) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::Percentage);
    v.part_number().decimal_places(decimal).push();
    v.part_text("%");
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
    v.part_currency()
        .locale(symbol_locale)
        .symbol(symbol.into())
        .push();
    v.part_text(" ");
    v.part_number()
        .decimal_places(2)
        .min_decimal_places(2)
        .grouping()
        .push();
    v.part_number()
        .decimal_places(2)
        .min_decimal_places(2)
        .grouping()
        .push();
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
    v.part_number()
        .decimal_places(2)
        .min_decimal_places(2)
        .grouping()
        .push();
    v.part_text(" ");
    v.part_currency()
        .locale(symbol_locale)
        .symbol(symbol.into())
        .push();
    v
}

/// Creates a new date format D.M.Y
pub fn create_loc_date_dmy_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::DateTime);
    v.part_day().style(FormatNumberStyle::Long).push();
    v.part_text(".");
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text(".");
    v.part_year().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a new date format M/D/Y
pub fn create_loc_date_mdy_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::DateTime);
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text("/");
    v.part_day().style(FormatNumberStyle::Long).push();
    v.part_text("/");
    v.part_year().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a datetime format Y-M-D H:M:S
pub fn create_loc_datetime_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::DateTime);
    v.part_day().style(FormatNumberStyle::Long).push();
    v.part_text(".");
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text(".");
    v.part_year().style(FormatNumberStyle::Long).push();
    v.part_text(" ");
    v.part_hours().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_minutes().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_seconds().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a new time-Duration format H:M:S
pub fn create_loc_time_format<S: Into<String>>(name: S, locale: Locale) -> ValueFormat {
    let mut v = ValueFormat::new_localized(name.into(), locale, ValueType::TimeDuration);
    v.part_hours().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_minutes().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_seconds().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a new number format.
pub fn create_boolean_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Boolean);
    v.part_boolean();
    v
}

/// Creates a new number format.
pub fn create_number_format<S: Into<String>>(name: S, decimal: u8, grouping: bool) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Number);
    v.part_number()
        .decimal_places(decimal)
        .test(grouping, |p| p.grouping())
        .push();
    v
}

/// Creates a new number format with a fixed number of decimal places.
pub fn create_number_format_fixed<S: Into<String>>(
    name: S,
    decimal: u8,
    grouping: bool,
) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Number);
    v.part_number()
        .fixed_decimal_places(decimal)
        .test(grouping, |p| p.grouping())
        .push();
    v
}

/// Creates a new percentage format.
pub fn create_percentage_format<S: Into<String>>(name: S, decimal: u8) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::Percentage);
    v.part_number().fixed_decimal_places(decimal).push();
    v.part_text("%");
    v
}

/// Creates a new currency format.
pub fn create_currency_prefix<S1, S2>(name: S1, symbol_locale: Locale, symbol: S2) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
{
    let mut v = ValueFormat::new_named(name.into(), ValueType::Currency);
    v.part_currency()
        .locale(symbol_locale)
        .symbol(symbol.into())
        .push();
    v.part_text(" ");
    v.part_number().fixed_decimal_places(2).grouping().push();
    v
}

/// Creates a new currency format.
pub fn create_currency_suffix<S1, S2>(name: S1, symbol_locale: Locale, symbol: S2) -> ValueFormat
where
    S1: Into<String>,
    S2: Into<String>,
{
    let mut v = ValueFormat::new_named(name.into(), ValueType::Currency);
    v.part_number().fixed_decimal_places(2).grouping().push();
    v.part_text(" ");
    v.part_currency()
        .locale(symbol_locale)
        .symbol(symbol.into())
        .push();
    v
}

/// Creates a new date format YYYY-MM-DD
pub fn create_date_iso_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.part_year().style(FormatNumberStyle::Long).push();
    v.part_text("-");
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text("-");
    v.part_day().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a new date format D.M.Y
pub fn create_date_dmy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.part_day().style(FormatNumberStyle::Long).push();
    v.part_text(".");
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text(".");
    v.part_year().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a new date format M/D/Y
pub fn create_date_mdy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text("/");
    v.part_day().style(FormatNumberStyle::Long).push();
    v.part_text("/");
    v.part_year().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a datetime format Y-M-D H:M:S
pub fn create_datetime_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::DateTime);
    v.part_year().style(FormatNumberStyle::Long).push();
    v.part_text("-");
    v.part_month().style(FormatNumberStyle::Long).push();
    v.part_text("-");
    v.part_day().style(FormatNumberStyle::Long).push();
    v.part_text(" ");
    v.part_hours().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_minutes().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_seconds().style(FormatNumberStyle::Long).push();
    v
}

/// Creates a new time-Duration format H:M:S
pub fn create_time_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::new_named(name.into(), ValueType::TimeDuration);
    v.part_hours().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_minutes().style(FormatNumberStyle::Long).push();
    v.part_text(":");
    v.part_seconds().style(FormatNumberStyle::Long).push();
    v
}
