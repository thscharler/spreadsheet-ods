use spreadsheet_ods::io::read_ods;
use spreadsheet_ods::OdsError;

#[test]
fn pagelayout() -> Result<(), OdsError> {
    let ods = read_ods("test_out/experiment.ods")?;

    println!("{:?}", ods);

    Ok(())
}