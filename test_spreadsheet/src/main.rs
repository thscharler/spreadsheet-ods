use spreadsheet_ods::{OdsError, read_ods, write_ods};

fn main() {
    run().unwrap();
}

fn run() -> Result<(), OdsError> {

    let wb = read_ods("../tests/missing/RE20010.ods")?;
    write_ods(&wb, "../test_out/RE20010.ods")?;

    // println!("wb {:?}", wb);

    Ok(())
}
