use chrono::NaiveDateTime;

use spreadsheet_ods::format::{FormatCalendarStyle, FormatNumberStyle};
use spreadsheet_ods::style::CellStyle;
use spreadsheet_ods::{write_ods, OdsError, Sheet, ValueFormat, ValueType, WorkBook};

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

    let mut v1 = ValueFormat::new_with_name("f1", ValueType::Number);
    v1.push_scientific(4);
    let v1 = wb.add_format(v1);

    let mut v2 = ValueFormat::new_with_name("f2", ValueType::Number);
    v2.push_number_fix(2, false);
    let v2 = wb.add_format(v2);

    let mut v3 = ValueFormat::new_with_name("f3", ValueType::Number);
    v3.push_number(2, false);
    let v3 = wb.add_format(v3);

    let mut v31 = ValueFormat::new_with_name("f31", ValueType::Number);
    v31.push_fraction(13, 1, 1, 1, false);
    let v31 = wb.add_format(v31);

    let mut v4 = ValueFormat::new_with_name("f4", ValueType::Currency);
    v4.push_currency("AT", "de", "€");
    v4.push_text(" ");
    v4.push_number(2, false);
    let v4 = wb.add_format(v4);

    let mut v5 = ValueFormat::new_with_name("f5", ValueType::Percentage);
    v5.push_number(2, false);
    v5.push_text("/ct");
    let v5 = wb.add_format(v5);

    let mut v6 = ValueFormat::new_with_name("f6", ValueType::Boolean);
    v6.push_boolean();
    let v6 = wb.add_format(v6);

    let mut v7 = ValueFormat::new_with_name("f7", ValueType::DateTime);
    v7.push_era(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    v7.push_text(" ");
    v7.push_year(FormatNumberStyle::Long);
    v7.push_text(" ");
    v7.push_month(FormatNumberStyle::Long);
    v7.push_text(" ");
    v7.push_day(FormatNumberStyle::Long);
    v7.push_text(" ");
    v7.push_day_of_week(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    v7.push_text(" ");
    v7.push_week_of_year(FormatCalendarStyle::Gregorian);
    v7.push_text(" ");
    v7.push_quarter(FormatNumberStyle::Long, FormatCalendarStyle::Gregorian);
    let v7 = wb.add_format(v7);

    let f1 = wb.add_cell_style(CellStyle::new("f1", &v1));
    let f2 = wb.add_cell_style(CellStyle::new("f2", &v2));
    let f3 = wb.add_cell_style(CellStyle::new("f3", &v3));
    let f31 = wb.add_cell_style(CellStyle::new("f31", &v31));
    let f4 = wb.add_cell_style(CellStyle::new("f4", &v4));
    let f5 = wb.add_cell_style(CellStyle::new("f5", &v5));
    let f6 = wb.add_cell_style(CellStyle::new("f6", &v6));
    let f7 = wb.add_cell_style(CellStyle::new("f7", &v7));

    let mut sh = Sheet::new();
    sh.set_styled_value(0, 0, 1.234567f64, &f1);
    sh.set_styled_value(1, 0, 1.234567f64, &f2);
    sh.set_styled_value(2, 0, 1.234567f64, &f3);
    sh.set_styled_value(2, 1, 1.234567f64, &f31);
    sh.set_styled_value(3, 0, 1.234567f64, &f4);
    sh.set_styled_value(4, 0, 1.234567f64, &f5);

    sh.set_styled_value(6, 0, 1.234567f64, &f6);

    sh.set_styled_value(
        7,
        0,
        NaiveDateTime::from_timestamp(1_223_222_222, 22992),
        &f7,
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
