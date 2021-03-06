use crate::format::ValueFormatRef;
use crate::style::CellStyle;
use crate::{format, CellStyleRef, ValueType, WorkBook};

///
/// Allows access to the value-format names for the default formats
/// as created by create_default_styles.
///
pub struct DefaultFormat {}

impl DefaultFormat {
    pub fn default() -> ValueFormatRef {
        ValueFormatRef::from("")
    }

    pub fn bool() -> ValueFormatRef {
        ValueFormatRef::from("bool1")
    }

    pub fn num() -> ValueFormatRef {
        ValueFormatRef::from("num1")
    }

    pub fn percent() -> ValueFormatRef {
        ValueFormatRef::from("percent1")
    }

    pub fn currency() -> ValueFormatRef {
        ValueFormatRef::from("currency1")
    }

    pub fn date() -> ValueFormatRef {
        ValueFormatRef::from("date1")
    }

    pub fn datetime() -> ValueFormatRef {
        ValueFormatRef::from("datetime1")
    }

    pub fn time() -> ValueFormatRef {
        ValueFormatRef::from("time1")
    }
}

///
/// Allows access to the names of the default styles as created by
/// create_default_styles.
///
pub struct DefaultStyle {}

impl DefaultStyle {
    pub fn bool() -> CellStyleRef {
        CellStyleRef::from("default-bool")
    }

    pub fn num() -> CellStyleRef {
        CellStyleRef::from("default-num")
    }

    pub fn percent() -> CellStyleRef {
        CellStyleRef::from("default-percent")
    }

    pub fn currency() -> CellStyleRef {
        CellStyleRef::from("default-currency")
    }

    pub fn date() -> CellStyleRef {
        CellStyleRef::from("default-date")
    }

    pub fn datetime() -> CellStyleRef {
        CellStyleRef::from("default-datetime")
    }

    pub fn time() -> CellStyleRef {
        CellStyleRef::from("default-time")
    }
}

/// Adds default-styles for all basic ValueTypes. These are also set as default
/// styles for the respective types. By calling this function for a new workbook,
/// the basic formatting is done.
///
/// This function is best seen as an example, as there is currently now
/// I18N support. So I set this up as it suited me.
///
pub fn create_default_styles(book: &mut WorkBook) {
    book.add_format(format::create_boolean_format(
        DefaultFormat::bool().to_string(),
    ));
    book.add_format(format::create_number_format(
        DefaultFormat::num().to_string(),
        2,
        false,
    ));
    book.add_format(format::create_percentage_format(
        DefaultFormat::percent().to_string(),
        2,
    ));
    book.add_format(format::create_currency_prefix(
        DefaultFormat::currency().to_string(),
        "de",
        "AT",
        "€",
    ));
    book.add_format(format::create_date_dmy_format(
        DefaultFormat::date().to_string(),
    ));
    book.add_format(format::create_datetime_format(
        DefaultFormat::datetime().to_string(),
    ));
    book.add_format(format::create_time_format(
        DefaultFormat::time().to_string(),
    ));

    book.add_cellstyle(CellStyle::new(
        DefaultStyle::bool().to_string(),
        &DefaultFormat::bool(),
    ));
    book.add_cellstyle(CellStyle::new(
        DefaultStyle::num().to_string(),
        &DefaultFormat::num(),
    ));
    book.add_cellstyle(CellStyle::new(
        DefaultStyle::percent().to_string(),
        &DefaultFormat::percent(),
    ));
    book.add_cellstyle(CellStyle::new(
        DefaultStyle::currency().to_string(),
        &DefaultFormat::currency(),
    ));
    book.add_cellstyle(CellStyle::new(
        DefaultStyle::date().to_string(),
        &DefaultFormat::date(),
    ));
    book.add_cellstyle(CellStyle::new(
        DefaultStyle::datetime().to_string(),
        &DefaultFormat::datetime(),
    ));
    book.add_cellstyle(CellStyle::new(
        DefaultStyle::time().to_string(),
        &DefaultFormat::time(),
    ));

    book.add_def_style(ValueType::Boolean, &DefaultStyle::bool());
    book.add_def_style(ValueType::Number, &DefaultStyle::num());
    book.add_def_style(ValueType::Percentage, &DefaultStyle::percent());
    book.add_def_style(ValueType::Currency, &DefaultStyle::currency());
    book.add_def_style(ValueType::DateTime, &DefaultStyle::date());
    book.add_def_style(ValueType::TimeDuration, &DefaultStyle::time());
}
