//!
//! Creates default formats for a new Workbook.
//!

use crate::format::ValueFormatRef;
use crate::{CellStyleRef, WorkBook};

///
/// Allows access to the value-format names for the default formats
/// as created by create_default_styles.
///
#[derive(Debug)]
pub struct DefaultFormat {}

impl DefaultFormat {
    /// Default format.
    #[allow(clippy::should_implement_trait)]
    pub fn default() -> ValueFormatRef {
        ValueFormatRef::from("")
    }

    /// Default boolean format.
    pub fn bool() -> ValueFormatRef {
        ValueFormatRef::from("bool1")
    }

    /// Default number format.
    pub fn number() -> ValueFormatRef {
        ValueFormatRef::from("num1")
    }

    /// Default percentage format.
    pub fn percent() -> ValueFormatRef {
        ValueFormatRef::from("percent1")
    }

    /// Default currency format.
    pub fn currency() -> ValueFormatRef {
        ValueFormatRef::from("currency1")
    }

    /// Default date format.
    pub fn date() -> ValueFormatRef {
        ValueFormatRef::from("date1")
    }

    /// Default datetime format.
    pub fn datetime() -> ValueFormatRef {
        ValueFormatRef::from("datetime1")
    }

    /// Default time format.
    pub fn time_of_day() -> ValueFormatRef {
        ValueFormatRef::from("time1")
    }

    /// Default time format.
    pub fn time_interval() -> ValueFormatRef {
        ValueFormatRef::from("interval1")
    }
}

///
/// Allows access to the names of the default styles as created by
/// create_default_styles.
///
#[derive(Debug)]
pub struct DefaultStyle {}

impl DefaultStyle {
    pub const BOOL: &'static str = "default-bool";
    pub const NUMBER: &'static str = "default-bool";
    pub const PERCENT: &'static str = "default-bool";
    pub const CURRENCY: &'static str = "default-bool";
    pub const DATE: &'static str = "default-bool";
    pub const DATETIME: &'static str = "default-bool";
    pub const TIME_OF_DAY: &'static str = "default-bool";
    pub const TIME_INTERVAL: &'static str = "default-bool";

    /// Default bool style.
    pub fn bool(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::BOOL).expect("style")
    }

    /// Default number style.
    pub fn number(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::NUMBER).expect("style")
    }

    /// Default percent style.
    pub fn percent(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::PERCENT).expect("style")
    }

    /// Default currency style.
    pub fn currency(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::CURRENCY).expect("style")
    }

    /// Default date style.
    pub fn date(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::DATE).expect("style")
    }

    /// Default datetime style.
    pub fn datetime(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::DATE).expect("style")
    }

    /// Default time style.
    pub fn time_of_day(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::TIME_OF_DAY).expect("style")
    }

    /// Default time style.
    pub fn time_interval(wb: &WorkBook) -> CellStyleRef {
        wb.cellstyle_ref(DefaultStyle::TIME_INTERVAL)
            .expect("style")
    }
}
//
// /// Replaced with WorkBook::locale_settings() or WorkBook::new(l: Locale).
// #[deprecated]
// pub fn create_default_styles(book: &mut WorkBook) {
//     book.add_boolean_format(format::create_boolean_format(
//         DefaultFormat::bool().to_string(),
//     ));
//     book.add_number_format(format::create_number_format(
//         DefaultFormat::number().to_string(),
//         2,
//         false,
//     ));
//     book.add_percentage_format(format::create_percentage_format(
//         DefaultFormat::percent().to_string(),
//         2,
//     ));
//     book.add_currency_format(format::create_currency_prefix(
//         DefaultFormat::currency().to_string(),
//         locale!("de_AT"),
//         "€",
//     ));
//     book.add_datetime_format(format::create_date_dmy_format(
//         DefaultFormat::date().to_string(),
//     ));
//     book.add_datetime_format(format::create_datetime_format(
//         DefaultFormat::datetime().to_string(),
//     ));
//     book.add_timeduration_format(format::create_time_of_day_format(
//         DefaultFormat::time_of_day().to_string(),
//     ));
//     book.add_timeduration_format(format::create_time_interval_format(
//         DefaultFormat::time_interval().to_string(),
//     ));
//
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::bool().to_string(),
//         &DefaultFormat::bool(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::number().to_string(),
//         &DefaultFormat::number(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::percent().to_string(),
//         &DefaultFormat::percent(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::currency().to_string(),
//         &DefaultFormat::currency(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::date().to_string(),
//         &DefaultFormat::date(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::datetime().to_string(),
//         &DefaultFormat::datetime(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::time_of_day().to_string(),
//         &DefaultFormat::time_of_day(),
//     ));
//     book.add_cellstyle(CellStyle::new(
//         DefaultStyle::time_interval().to_string(),
//         &DefaultFormat::time_interval(),
//     ));
//
//     book.add_def_style(ValueType::Boolean, &DefaultStyle::bool());
//     book.add_def_style(ValueType::Number, &DefaultStyle::number());
//     book.add_def_style(ValueType::Percentage, &DefaultStyle::percent());
//     book.add_def_style(ValueType::Currency, &DefaultStyle::currency());
//     book.add_def_style(ValueType::DateTime, &DefaultStyle::date());
//     book.add_def_style(ValueType::TimeDuration, &DefaultStyle::time_interval());
// }
