use crate::defaultstyles::DefaultFormat;
use crate::format::{
    create_loc_boolean_format, create_loc_currency_prefix, create_loc_date_mdy_format,
    create_loc_datetime_format, create_loc_number_format, create_loc_percentage_format,
    FormatNumberStyle,
};
use crate::locale::LocalizedValueFormat;
use crate::{ValueFormat, ValueType};
use icu_locid::{locale, Locale};

pub(crate) struct LocaleEnUs {}

pub(crate) static LOCALE_EN_US: LocaleEnUs = LocaleEnUs {};

impl LocaleEnUs {
    const LOCALE: Locale = locale!("en_US");
}

impl LocalizedValueFormat for LocaleEnUs {
    fn locale(&self) -> Locale {
        LocaleEnUs::LOCALE
    }

    fn boolean_format(&self) -> ValueFormat {
        create_loc_boolean_format(DefaultFormat::bool(), LocaleEnUs::LOCALE)
    }

    fn number_format(&self) -> ValueFormat {
        create_loc_number_format(DefaultFormat::num(), LocaleEnUs::LOCALE, 2, false)
    }

    fn percentage_format(&self) -> ValueFormat {
        create_loc_percentage_format(DefaultFormat::percent(), LocaleEnUs::LOCALE, 2)
    }

    fn currency_format(&self) -> ValueFormat {
        create_loc_currency_prefix(
            DefaultFormat::currency(),
            LocaleEnUs::LOCALE,
            LocaleEnUs::LOCALE,
            "$",
        )
    }

    fn date_format(&self) -> ValueFormat {
        create_loc_date_mdy_format(DefaultFormat::date(), LocaleEnUs::LOCALE)
    }

    fn datetime_format(&self) -> ValueFormat {
        create_loc_datetime_format(DefaultFormat::datetime(), LocaleEnUs::LOCALE)
    }

    fn time_format(&self) -> ValueFormat {
        let mut v = ValueFormat::new_localized(
            DefaultFormat::time(),
            LocaleEnUs::LOCALE,
            ValueType::TimeDuration,
        );
        v.part_hours().style(FormatNumberStyle::Long).push();
        v.part_text(":");
        v.part_minutes().style(FormatNumberStyle::Long).push();
        v.part_text(":");
        v.part_seconds().style(FormatNumberStyle::Long).push();
        v.part_text(" ");
        v.part_am_pm();
        v
    }
}
