use chrono::NaiveDateTime;

use spreadsheet_ods::{OdsError, Sheet, Style, ValueFormat, ValueType, WorkBook, write_ods};
use spreadsheet_ods::format::{FormatCalendarStyle, FormatNumberStyle};

#[test]
pub fn value_format() {
    let mut f0 = ValueFormat::new();
    f0.push_boolean();
    assert_eq!(f0.format_boolean(true), "true");
    assert_eq!(f0.format_float(1f64), "");

    let mut f1 = ValueFormat::new();
    f1.push_number(3, true);
    assert_eq!(f1.format_boolean(true), "");
    // these are questionable ...
    // but i wrote somewhere there is no i18n support yet, so ...
    // todo: should be '1,234'
    assert_eq!(f1.format_float(1.2345f64), "1.234");
    // todo: should be '1,2'
    assert_eq!(f1.format_float(1.2f64), "1.200");

    let mut f2 = ValueFormat::new();
    f2.push_currency("AT", "de", "€");
    f2.push_number_fix(2, true);
    // todo: should be '€ 1,33'
    assert_eq!(f2.format_float(1.333f64), "€1.33");

    let mut f3 = ValueFormat::new();
    f3.push_fraction(10, 1, 1, 1, false);
    // todo: should be '1 32/10' or the like
    assert_eq!(f3.format_float(1.3223f64), "");

    let mut f4 = ValueFormat::new();
    f4.push_scientific(5);
    // todo: should be '3.12345e0'
    assert_eq!(f4.format_float(3.123456), "3.123456e0");

    let mut f5 = ValueFormat::new();
    f5.push_era(FormatNumberStyle::Short, FormatCalendarStyle::Gregorian);
    f5.push_text(" ");
    f5.push_day(FormatNumberStyle::Short);
    f5.push_text(" ");
    f5.push_month(FormatNumberStyle::Long);
    f5.push_text(" ");
    f5.push_year(FormatNumberStyle::Long);
    // todo: should be 'AD 12 02 2009'
    assert_eq!(
        f5.format_datetime(&NaiveDateTime::from_timestamp(1234442333, 12234332)),
        " 12 02 2009"
    );

    let mut f6 = ValueFormat::new();
    f6.push_day_of_week(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    assert_eq!(
        f6.format_datetime(&NaiveDateTime::from_timestamp(1234442333, 12234332)),
        "Thursday"
    );

    let mut f7 = ValueFormat::new();
    f7.push_week_of_year(FormatCalendarStyle::Gregorian);
    assert_eq!(
        f7.format_datetime(&NaiveDateTime::from_timestamp(1234442333, 12234332)),
        "6"
    );

    let mut f8 = ValueFormat::new();
    f8.push_quarter(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    // todo: ???
    assert_eq!(
        f8.format_datetime(&NaiveDateTime::from_timestamp(1234442333, 12234332)),
        ""
    );

    let mut f9 = ValueFormat::new();
    f9.push_hours(FormatNumberStyle::Long);
    f9.push_minutes(FormatNumberStyle::Long);
    f9.push_seconds(FormatNumberStyle::Long);
    assert_eq!(
        f9.format_datetime(&NaiveDateTime::from_timestamp(1234442333, 12234332)),
        "123853"
    );
}

#[test]
fn write_format() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut fs = ValueFormat::new_with_name("f1", ValueType::Number);
    fs.push_scientific(4);
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f2", ValueType::Number);
    fs.push_number_fix(2, false);
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f3", ValueType::Number);
    fs.push_number(2, false);
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f31", ValueType::Number);
    fs.push_fraction(13, 1, 1, 1, false);
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f4", ValueType::Currency);
    fs.push_currency("AT", "de", "€");
    fs.push_text(" ");
    fs.push_number(2, false);
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f5", ValueType::Percentage);
    fs.push_number(2, false);
    fs.push_text("/ct");
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f6", ValueType::Boolean);
    fs.push_boolean();
    wb.add_format(fs);

    let mut fs = ValueFormat::new_with_name("f7", ValueType::DateTime);
    fs.push_era(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    fs.push_text(" ");
    fs.push_year(FormatNumberStyle::Long);
    fs.push_text(" ");
    fs.push_month(FormatNumberStyle::Long);
    fs.push_text(" ");
    fs.push_day(FormatNumberStyle::Long);
    fs.push_text(" ");
    fs.push_day_of_week(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    fs.push_text(" ");
    fs.push_week_of_year(FormatCalendarStyle::Gregorian);
    fs.push_text(" ");
    fs.push_quarter(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    wb.add_format(fs);

    wb.add_style(Style::new_cell_style("f1", "f1"));
    wb.add_style(Style::new_cell_style("f2", "f2"));
    wb.add_style(Style::new_cell_style("f3", "f3"));
    wb.add_style(Style::new_cell_style("f31", "f31"));
    wb.add_style(Style::new_cell_style("f4", "f4"));
    wb.add_style(Style::new_cell_style("f5", "f5"));
    wb.add_style(Style::new_cell_style("f6", "f6"));
    wb.add_style(Style::new_cell_style("f7", "f7"));

    let mut sh = Sheet::new();
    sh.set_styled_value(0, 0, 1.234567f64, "f1");
    sh.set_styled_value(1, 0, 1.234567f64, "f2");
    sh.set_styled_value(2, 0, 1.234567f64, "f3");
    sh.set_styled_value(2, 1, 1.234567f64, "f31");
    sh.set_styled_value(3, 0, 1.234567f64, "f4");
    sh.set_styled_value(4, 0, 1.234567f64, "f5");

    sh.set_styled_value(6, 0, 1.234567f64, "f6");

    sh.set_styled_value(
        7,
        0,
        NaiveDateTime::from_timestamp(1_223_222_222, 22992),
        "f7",
    );

    wb.push_sheet(sh);
    let path = std::path::Path::new("test_out/format.ods");
    if path.exists() {
        write_ods(&wb, path)
    } else {
        std::fs::create_dir_all(path.parent().unwrap())?;
        std::fs::File::create(path)?;
        write_ods(&wb, path)
    }
}
