use std::path::Path;

use spreadsheet_ods::{read_ods, write_ods, OdsError, Sheet, SheetSplitMode, ValueType, WorkBook};

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

    wb.config_mut().has_sheet_tabs = false;

    let cc = wb.sheet_mut(0).config_mut();
    cc.show_grid = true;
    cc.vert_split_pos = 2;
    cc.vert_split_mode = SheetSplitMode::Cell;

    write_ods(&mut wb, "test_out/orders.ods")?;
    Ok(())
}

#[test]
fn test_write_read_write_read() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();
    let mut sh = Sheet::new();

    sh.set_value(0, 0, "A");
    wb.push_sheet(sh);

    let path = Path::new("tests/rw.ods");
    let temp = Path::new("test_out/rw.ods");

    std::fs::copy(path, temp)?;

    let _ods = read_ods(temp)?;

    write_ods(&mut wb, temp)?;

    let _ods = read_ods(temp)?;

    Ok(())
}
