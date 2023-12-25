use spreadsheet_ods::{OdsError, WorkBook};
use std::fs;
use std::path::Path;

#[allow(dead_code)]
pub fn test_write_ods<P: AsRef<Path>>(book: &mut WorkBook, ods_path: P) -> Result<(), OdsError> {
    fs::create_dir_all("test_out")?;
    spreadsheet_ods::write_ods(book, ods_path)
}
