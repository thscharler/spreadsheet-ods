mod lib_test;

use lib_test::*;
use spreadsheet_ods::{read_ods, write_ods, OdsError};

#[test]
fn test_draw_image() -> Result<(), OdsError> {
    let mut wb = read_ods("tests/draw_image.ods")?;

    let sh = wb.sheet(0);
    assert!(sh.draw_frames(1, 1).is_some());

    test_write_ods(&mut wb, "test_out/draw_image.ods")?;
    let wb = read_ods("test_out/draw_image.ods")?;

    let sh = wb.sheet(0);
    assert!(sh.draw_frames(1, 1).is_some());

    Ok(())
}
