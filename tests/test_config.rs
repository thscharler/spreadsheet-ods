use spreadsheet_ods::{read_ods, write_ods, OdsError, SplitMode};

#[test]
fn read_orders() -> Result<(), OdsError> {
    let mut wb = read_ods("tests/orders.ods")?;

    wb.config_mut().has_sheet_tabs = false;

    let cc = wb.sheet_mut(0).config_mut();
    cc.show_grid = true;
    cc.vert_split_pos = 2;
    cc.vert_split_mode = SplitMode::Heading;

    write_ods(&mut wb, "test_out/orders.ods")?;
    Ok(())
}
