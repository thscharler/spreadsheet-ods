use spreadsheet_ods::{OdsError, read_ods, write_ods};

fn main() {
    run().unwrap();
}

fn run() -> Result<(), OdsError> {

    let wb = read_ods("../tests/missing/2017 Rindfleisch.ods")?;
    write_ods(&wb, "../test_out/2017 Rindfleisch.ods")?;

    //println!("wb {:?}", wb);

    Ok(())
}
