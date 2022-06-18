//!
//! Defines localized versions for all default formats.
//!

#[cfg(feature = "locale_de_AT")]
mod de_at;
#[cfg(feature = "locale_en_US")]
mod en_us;
mod und;

use crate::ValueFormat;
use icu_locid::Locale;
use lazy_static::lazy_static;
use std::collections::HashMap;

/// Defines functions that generate the standard formats for various
/// value types.
pub trait LocalizedValueFormat: Sync {
    /// Default boolean format.
    fn boolean_format(&self) -> ValueFormat;
    /// Default number format.
    fn number_format(&self) -> ValueFormat;
    /// Default percentage format.
    fn percentage_format(&self) -> ValueFormat;
    /// Default currency format.
    fn currency_format(&self) -> ValueFormat;
    /// Default date format.
    fn date_format(&self) -> ValueFormat;
    /// Default date/time format.
    fn datetime_format(&self) -> ValueFormat;
    /// Default time format.
    fn time_format(&self) -> ValueFormat;
}

lazy_static! {
    static ref LOCALE_DATA: HashMap<Locale, &'static dyn LocalizedValueFormat> = {
        #[allow(unused_mut)]
        let mut lm: HashMap<Locale, &'static dyn LocalizedValueFormat> = HashMap::new();

        #[cfg(feature = "locale_de_AT")]
        {
            lm.insert(locale!("de_AT"), &de_at::LOCALE_DE_AT);
        }
        #[cfg(feature = "locale_en_US")]
        {
            lm.insert(locale!("en_US"), &en_us::LOCALE_EN_US);
        }
        lm
    };
}

/// Returns the localized format or a fallback.
pub fn localized_format(locale: Locale) -> &'static dyn LocalizedValueFormat {
    match LOCALE_DATA.get(&locale) {
        None => &und::LOCALE_UND,
        Some(loc) => *loc,
    }
}
