use crate::{format, ValueType, WorkBook};
use crate::style::Style;

/// Adds default-styles for all basic ValueTypes. These are also set as default
/// styles for the respective types. By calling this function for a new workbook,
/// the basic formatting is done.
///
/// This function is best seen as an example, as there is currently now
/// I18N support. So I set this up as it suited me.
///
pub fn create_default_styles(book: &mut WorkBook) {
    book.add_format(format::create_boolean_format("bool1"));
    book.add_format(format::create_number_format("num1", 2, false));
    book.add_format(format::create_percentage_format("percent1", 2));
    book.add_format(format::create_currency_prefix("currency1", "de", "AT", "€"));
    book.add_format(format::create_date_dmy_format("date1"));
    book.add_format(format::create_datetime_format("datetime1"));
    book.add_format(format::create_time_format("time1"));

    book.add_style(Style::cell_style("default-bool", "bool1"));
    book.add_style(Style::cell_style("default-num", "num1"));
    book.add_style(Style::cell_style("default-percent", "percent1"));
    book.add_style(Style::cell_style("default-currency", "currency1"));
    book.add_style(Style::cell_style("default-date", "date1"));
    book.add_style(Style::cell_style("default-time", "time1"));

    book.add_def_style(ValueType::Boolean, "default-bool");
    book.add_def_style(ValueType::Number, "default-num");
    book.add_def_style(ValueType::Percentage, "default-percent");
    book.add_def_style(ValueType::Currency, "default-currency");
    book.add_def_style(ValueType::DateTime, "default-date");
    book.add_def_style(ValueType::TimeDuration, "default-time");
}
