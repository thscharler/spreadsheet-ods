use spreadsheet_ods::{read_fods, read_ods, write_fods, OdsError};

#[test]
fn read_write_fods() -> Result<(), OdsError> {
    let mut wb = read_ods("tests/orders.ods")?;
    write_fods(&mut wb, "test_out/orders.fods")?;
    let _wb = read_fods("test_out/orders.fods")?;
    Ok(())
}
