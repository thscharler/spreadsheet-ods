use spreadsheet_ods::io::{read_ods, write_ods};
use spreadsheet_ods::OdsError;

#[test]
fn pagelayout() -> Result<(), OdsError> {
    let ods = read_ods("test_out/experiment.ods")?;
    write_ods(&ods, "test_out/rexp.ods")?;

    Ok(())
}