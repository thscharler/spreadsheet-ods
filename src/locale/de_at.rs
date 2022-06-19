use crate::defaultstyles::DefaultFormat;
use crate::format::{
    create_loc_boolean_format, create_loc_currency_prefix, create_loc_date_dmy_format,
    create_loc_datetime_format, create_loc_number_format, create_loc_percentage_format,
    create_loc_time_format,
};
use crate::locale::LocalizedValueFormat;
use crate::ValueFormat;
use icu_locid::{locale, Locale};

pub(crate) struct LocaleDeAt {}

pub(crate) static LOCALE_DE_AT: LocaleDeAt = LocaleDeAt {};

impl LocaleDeAt {
    const LOCALE: Locale = locale!("de_AT");
}

impl LocalizedValueFormat for LocaleDeAt {
    fn locale(&self) -> Locale {
        LocaleDeAt::LOCALE
    }

    fn boolean_format(&self) -> ValueFormat {
        create_loc_boolean_format(DefaultFormat::bool(), LocaleDeAt::LOCALE)
    }

    fn number_format(&self) -> ValueFormat {
        create_loc_number_format(DefaultFormat::num(), LocaleDeAt::LOCALE, 2, false)
    }

    fn percentage_format(&self) -> ValueFormat {
        create_loc_percentage_format(DefaultFormat::percent(), LocaleDeAt::LOCALE, 2)
    }

    fn currency_format(&self) -> ValueFormat {
        create_loc_currency_prefix(
            DefaultFormat::currency(),
            LocaleDeAt::LOCALE,
            LocaleDeAt::LOCALE,
            "€",
        )
    }

    fn date_format(&self) -> ValueFormat {
        create_loc_date_dmy_format(DefaultFormat::date(), LocaleDeAt::LOCALE)
    }

    fn datetime_format(&self) -> ValueFormat {
        create_loc_datetime_format(DefaultFormat::datetime(), LocaleDeAt::LOCALE)
    }

    fn time_format(&self) -> ValueFormat {
        create_loc_time_format(DefaultFormat::datetime(), LocaleDeAt::LOCALE)
    }
}
