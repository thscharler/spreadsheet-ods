use icu_locid::{locale, Locale};
use spreadsheet_ods::{
    read_ods, write_ods, CellStyle, OdsError, Sheet, ValueFormat, ValueType, WorkBook,
};

#[test]
pub fn test_locale1() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_localized(locale!("de_AT"));
    let mut sheet = Sheet::new("sheet1");

    let mut v0 = ValueFormat::new_localized("v0", locale!("ru_RU"), ValueType::Currency);
    v0.push_number(2, true);
    v0.push_text(" ");
    v0.push_currency_symbol(locale!("ru_RU"), "");
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
    let mut wb = WorkBook::new_localized(Locale::UND);
    let mut sheet = Sheet::new("sheet1");

    sheet.set_value(1, 1, 1234);

    wb.push_sheet(sheet);

    write_ods(&mut wb, "test_out/locale2.ods")?;

    let _wb = read_ods("test_out/locale2.ods")?;

    Ok(())
}
