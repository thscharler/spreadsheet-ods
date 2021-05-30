use spreadsheet_ods::{read_ods, write_ods, OdsError, Sheet, ValueType, WorkBook};

#[test]
fn test_write_read() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();
    let mut sh = Sheet::new();

    sh.set_value(0, 0, "A");

    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/test_0.ods")?;

    let wi = read_ods("test_out/test_0.ods")?;
    let si = wi.sheet(0);

    assert_eq!(si.value(0, 0).as_str_or(""), "A");

    Ok(())
}

#[test]
fn read_text() -> Result<(), OdsError> {
    let wb = read_ods("tests/text.ods")?;
    let sh = wb.sheet(0);

    let v = sh.value(0, 0);

    assert_eq!(v.value_type(), ValueType::TextXml);

    Ok(())
}

#[test]
fn read_orders() -> Result<(), OdsError> {
    let mut wb = read_ods("tests/orders.ods")?;

    let cc = wb.sheet_mut(0).config_mut();
    cc.show_grid = false;
    cc.vert_split_pos = 3;
    cc.vert_split_mode = 1;

    write_ods(&mut wb, "test_out/orders.ods")?;
    Ok(())
}
