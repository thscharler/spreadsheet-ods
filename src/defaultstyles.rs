
use crate::{format, WorkBook, Style, Family, ValueType};

/// Adds default-styles for all basic ValueTypes. These are also set as default
/// styles for the respective types. By calling this function for a new workbook,
/// the basic formatting is done.
///
/// Beware
/// There is no i18n yet, so currency is set to euro for now.
/// And dates are european DMY style.
///
pub fn create_default_styles(book: &mut WorkBook) {
    book.add_format(format::create_boolean_format("BOOL1"));
    book.add_format(format::create_number_format("NUM1", 2, false));
    book.add_format(format::create_percentage_format("PERCENT1", 2));
    book.add_format(format::create_euro_format("CURRENCY1"));
    book.add_format(format::create_date_mdy_format("DATE1"));
    book.add_format(format::create_datetime_format("DATETIME1"));
    book.add_format(format::create_time_format("TIME1"));

    book.add_style(Style::with_name(Family::TableCell, "DEFAULT-BOOL", "BOOLEAN1"));
    book.add_style(Style::with_name(Family::TableCell, "DEFAULT-NUM", "NUM1"));
    book.add_style(Style::with_name(Family::TableCell, "DEFAULT-PERCENT", "PERCENT1"));
    book.add_style(Style::with_name(Family::TableCell, "DEFAULT-CURRENCY", "CURRENCY1"));
    book.add_style(Style::with_name(Family::TableCell, "DEFAULT-DATE", "DATE1"));
    book.add_style(Style::with_name(Family::TableCell, "DEFAULT-TIME", "TIME1"));

    book.add_def_style(ValueType::Boolean, "DEFAULT-BOOL");
    book.add_def_style(ValueType::Number, "DEFAULT-NUM");
    book.add_def_style(ValueType::Percentage, "DEFAULT-PERCENT");
    book.add_def_style(ValueType::Currency, "DEFAULT-CURRENCY");
    book.add_def_style(ValueType::DateTime, "DEFAULT-DATE");
    book.add_def_style(ValueType::TimeDuration, "DEFAULT-TIME");
}