use spreadsheet_ods::{read_ods, OdsError};

#[test]
fn read_orders() -> Result<(), OdsError> {
    let wb = read_ods("tests/Unbenannt 1.ods")?;

    dbg!(wb);

    Ok(())
}
