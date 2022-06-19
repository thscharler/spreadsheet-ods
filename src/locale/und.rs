use crate::defaultstyles::DefaultFormat;
use crate::format::{
    create_boolean_format, create_currency_prefix, create_date_iso_format, create_datetime_format,
    create_number_format, create_percentage_format, create_time_format,
};
use crate::locale::LocalizedValueFormat;
use crate::ValueFormat;
use icu_locid::{locale, Locale};

pub(crate) struct LocaleUnd {}

pub(crate) static LOCALE_UND: LocaleUnd = LocaleUnd {};

impl LocalizedValueFormat for LocaleUnd {
    fn locale(&self) -> Locale {
        Locale::UND
    }

    fn boolean_format(&self) -> ValueFormat {
        create_boolean_format(DefaultFormat::bool())
    }

    fn number_format(&self) -> ValueFormat {
        create_number_format(DefaultFormat::num(), 2, false)
    }

    fn percentage_format(&self) -> ValueFormat {
        create_percentage_format(DefaultFormat::percent(), 2)
    }

    fn currency_format(&self) -> ValueFormat {
        create_currency_prefix(DefaultFormat::currency(), locale!("en_US"), "")
    }

    fn date_format(&self) -> ValueFormat {
        create_date_iso_format(DefaultFormat::date())
    }

    fn datetime_format(&self) -> ValueFormat {
        create_datetime_format(DefaultFormat::datetime())
    }

    fn time_format(&self) -> ValueFormat {
        create_time_format(DefaultFormat::time())
    }
}
