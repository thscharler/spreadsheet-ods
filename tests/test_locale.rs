use chrono::{Duration, NaiveDate, NaiveDateTime};
use icu_locid::{locale, Locale};
use spreadsheet_ods::defaultstyles::DefaultStyle;
use spreadsheet_ods::{
    read_ods, write_ods, CellStyle, OdsError, Sheet, Value, ValueFormat, ValueType, WorkBook,
};

#[test]
pub fn test_locale1() -> Result<(), OdsError> {
    let mut wb = WorkBook::new(locale!("de_AT"));
    let mut sheet = Sheet::new("sheet1");

    let mut v0 = ValueFormat::new_localized("v0", locale!("ru_RU"), ValueType::Currency);
    v0.part_number().decimal_places(2).grouping().push();
    v0.part_text(" ");
    v0.part_currency().locale(locale!("ru_RU")).push();
    let v0 = wb.add_format(v0);

    let s0 = CellStyle::new("s0", v0.as_ref());
    let s0 = wb.add_cellstyle(s0);

    sheet.set_styled_value(1, 1, 47.11f64, s0.as_ref());

    wb.push_sheet(sheet);

    write_ods(&mut wb, "test_out/locale1.ods")?;

    Ok(())
}

#[test]
pub fn test_locale2() -> Result<(), OdsError> {
    let mut wb = WorkBook::new(Locale::UND);
    let mut sheet = Sheet::new("sheet1");

    sheet.set_styled_value(1, 1, 1234, &DefaultStyle::bool());
    sheet.set_styled_value(2, 1, 1234, &DefaultStyle::number());
    sheet.set_styled_value(3, 1, 1234, &DefaultStyle::percent());
    sheet.set_styled_value(4, 1, 1234, &DefaultStyle::currency());
    sheet.set_styled_value(5, 1, 1234, &DefaultStyle::date());
    sheet.set_styled_value(6, 1, 1234, &DefaultStyle::datetime());
    sheet.set_styled_value(7, 1, 1234, &DefaultStyle::time_of_day());
    sheet.set_styled_value(8, 1, 1234, &DefaultStyle::time_interval());

    wb.push_sheet(sheet);

    write_ods(&mut wb, "test_out/locale2.ods")?;

    let _wb = read_ods("test_out/locale2.ods")?;

    Ok(())
}

#[test]
pub fn test_locale3() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();
    let mut sheet = Sheet::new("sheet1");

    sheet.set_value(1, 1, Value::Boolean(true));
    sheet.set_value(2, 1, Value::Number(1234f64));
    sheet.set_value(3, 1, Value::Percentage(1234f64));
    sheet.set_value(4, 1, Value::new_currency("", 1234f64));
    sheet.set_value(
        5,
        1,
        Value::DateTime(NaiveDate::from_ymd(2000, 1, 1).and_hms(0, 0, 0)),
    );
    sheet.set_value(
        6,
        1,
        Value::DateTime(NaiveDateTime::from_timestamp(1234, 0)),
    );
    sheet.set_value(8, 1, Value::TimeDuration(Duration::hours(1234)));

    wb.push_sheet(sheet);

    write_ods(&mut wb, "test_out/locale3.ods")?;

    let _wb = read_ods("test_out/locale3.ods")?;

    Ok(())
}

#[test]
pub fn test_locale4() -> Result<(), OdsError> {
    let mut wb = WorkBook::new(locale!("en_GB"));
    let mut sheet = Sheet::new("sheet1");

    sheet.set_styled_value(1, 1, Value::Boolean(true), &DefaultStyle::bool());
    sheet.set_styled_value(2, 1, Value::Number(1234.5678f64), &DefaultStyle::number());
    sheet.set_styled_value(
        3,
        1,
        Value::Percentage(1234.5678f64),
        &DefaultStyle::percent(),
    );
    sheet.set_styled_value(
        4,
        1,
        Value::new_currency("GBP", 1234.5678f64),
        &DefaultStyle::currency(),
    );
    sheet.set_styled_value(
        5,
        1,
        Value::DateTime(NaiveDate::from_ymd(2000, 1, 1).and_hms(0, 0, 0)),
        &DefaultStyle::date(),
    );
    sheet.set_styled_value(
        6,
        1,
        Value::DateTime(NaiveDate::from_ymd(2000, 1, 1).and_hms(1, 2, 3)),
        &DefaultStyle::datetime(),
    );
    sheet.set_styled_value(
        7,
        1,
        Value::DateTime(NaiveDate::from_ymd(2000, 1, 1).and_hms(1, 2, 3)),
        &DefaultStyle::time_of_day(),
    );
    sheet.set_styled_value(
        8,
        1,
        Value::TimeDuration(Duration::hours(1234)),
        &DefaultStyle::time_interval(),
    );

    wb.push_sheet(sheet);

    write_ods(&mut wb, "test_out/locale4.ods")?;

    let _wb = read_ods("test_out/locale4.ods")?;

    Ok(())
}
