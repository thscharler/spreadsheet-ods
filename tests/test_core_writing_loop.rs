use spreadsheet_ods::{read_ods, write_ods, OdsError, Sheet, WorkBook};

// basic case, data in the very first row
#[test]
fn test_write_first() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(0, 0, 1);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/simple0.ods")?;

    let wb = read_ods("test_out/simple0.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(0, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, empty rows before the first data row.
#[test]
fn test_write_empty_before() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/simple1.ods")?;

    let wb = read_ods("test_out/simple1.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, row1 after row2
#[test]
fn test_write_simple() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(4, 0, 1);
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/simple2.ods")?;

    let wb = read_ods("test_out/simple2.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(4, 0).as_i32_or(0), 1);
    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, row1 gap row2
#[test]
fn test_write_gap() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(2, 0, 1);
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/simple3.ods")?;

    let wb = read_ods("test_out/simple3.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(2, 0).as_i32_or(0), 1);
    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, row1 gap row2
// row1 with a repeat of 2
#[test]
fn test_write_gap_repeat() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(2, 0, 1);
    sh.set_row_repeat(2, 2);
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/simple4.ods")?;

    let wb = read_ods("test_out/simple4.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(2, 0).as_i32_or(0), 1);
    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}
