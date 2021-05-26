use spreadsheet_ods::{cm, read_ods, write_ods, Length, OdsError, Sheet, WorkBook};

#[test]
fn test_colwidth() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut sh = Sheet::new();
    sh.set_value(0, 0, 1234);
    sh.set_col_width(0, cm!(2.54));
    sh.set_row_height(0, cm!(1.27));
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/colwidth.ods")?;

    let wb = read_ods("test_out/colwidth.ods")?;
    assert_eq!(wb.sheet(0).col_width(0), cm!(2.54));
    assert_eq!(wb.sheet(0).row_height(0), cm!(1.27));

    Ok(())
}
