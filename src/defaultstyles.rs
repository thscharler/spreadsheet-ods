use crate::format::ValueFormatRef;
use crate::style::CellStyle;
use crate::{format, CellStyleRef, ValueType, WorkBook};

#[macro_export]
macro_rules! value_bool {
    () => {
        ValueFormatRef::from("bool1")
    };
}

#[macro_export]
macro_rules! value_num {
    () => {
        ValueFormatRef::from("num1")
    };
}

#[macro_export]
macro_rules! value_percent {
    () => {
        ValueFormatRef::from("percent1")
    };
}

#[macro_export]
macro_rules! value_currency {
    () => {
        ValueFormatRef::from("currency1")
    };
}

#[macro_export]
macro_rules! value_date {
    () => {
        ValueFormatRef::from("date1")
    };
}

#[macro_export]
macro_rules! value_datetime {
    () => {
        ValueFormatRef::from("datetime1")
    };
}

#[macro_export]
macro_rules! value_time {
    () => {
        ValueFormatRef::from("time1")
    };
}

#[macro_export]
macro_rules! style_bool {
    () => {
        CellStyleRef::from("default-bool")
    };
}

#[macro_export]
macro_rules! style_num {
    () => {
        CellStyleRef::from("default-num")
    };
}

#[macro_export]
macro_rules! style_percent {
    () => {
        CellStyleRef::from("default-percent")
    };
}

#[macro_export]
macro_rules! style_currency {
    () => {
        CellStyleRef::from("default-currency")
    };
}

#[macro_export]
macro_rules! style_date {
    () => {
        CellStyleRef::from("default-date")
    };
}

#[macro_export]
macro_rules! style_datetime {
    () => {
        CellStyleRef::from("default-datetime")
    };
}

#[macro_export]
macro_rules! style_time {
    () => {
        CellStyleRef::from("default-time")
    };
}

/// Adds default-styles for all basic ValueTypes. These are also set as default
/// styles for the respective types. By calling this function for a new workbook,
/// the basic formatting is done.
///
/// This function is best seen as an example, as there is currently now
/// I18N support. So I set this up as it suited me.
///
pub fn create_default_styles(book: &mut WorkBook) {
    book.add_format(format::create_boolean_format(value_bool!().to_string()));
    book.add_format(format::create_number_format(
        value_num!().to_string(),
        2,
        false,
    ));
    book.add_format(format::create_percentage_format(
        value_percent!().to_string(),
        2,
    ));
    book.add_format(format::create_currency_prefix(
        value_currency!().to_string(),
        "de",
        "AT",
        "€",
    ));
    book.add_format(format::create_date_dmy_format(value_date!().to_string()));
    book.add_format(format::create_datetime_format(
        value_datetime!().to_string(),
    ));
    book.add_format(format::create_time_format(value_time!().to_string()));

    book.add_cell_style(CellStyle::new(style_bool!().to_string(), &value_bool!()));
    book.add_cell_style(CellStyle::new(style_num!().to_string(), &value_num!()));
    book.add_cell_style(CellStyle::new(
        style_percent!().to_string(),
        &value_percent!(),
    ));
    book.add_cell_style(CellStyle::new(
        style_currency!().to_string(),
        &value_currency!(),
    ));
    book.add_cell_style(CellStyle::new(style_date!().to_string(), &value_date!()));
    book.add_cell_style(CellStyle::new(
        style_datetime!().to_string(),
        &value_datetime!(),
    ));
    book.add_cell_style(CellStyle::new(style_time!().to_string(), &value_time!()));

    book.add_def_style(ValueType::Boolean, &style_bool!());
    book.add_def_style(ValueType::Number, &style_num!());
    book.add_def_style(ValueType::Percentage, &style_percent!());
    book.add_def_style(ValueType::Currency, &style_currency!());
    book.add_def_style(ValueType::DateTime, &style_date!());
    book.add_def_style(ValueType::TimeDuration, &style_time!());
}
