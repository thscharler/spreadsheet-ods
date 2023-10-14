mod lib_test;

use lib_test::*;
use spreadsheet_ods::{read_ods, write_ods, OdsError, OdsOptions, Sheet, WorkBook};
use std::fs;
use std::fs::File;
use std::io::BufReader;

// basic case, data in the very first row
#[test]
fn test_write_first() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(0, 0, 1);
    wb.push_sheet(sh);

    test_write_ods(&mut wb, "test_out/simple0.ods")?;

    let wb = read_ods("test_out/simple0.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(0, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, empty rows before the first data row.
#[test]
fn test_write_empty_before() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    test_write_ods(&mut wb, "test_out/simple1.ods")?;

    let wb = read_ods("test_out/simple1.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, row1 after row2
#[test]
fn test_write_simple() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(4, 0, 1);
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    test_write_ods(&mut wb, "test_out/simple2.ods")?;

    let wb = read_ods("test_out/simple2.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(4, 0).as_i32_or(0), 1);
    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, row1 gap row2
#[test]
fn test_write_gap() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(2, 0, 1);
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    test_write_ods(&mut wb, "test_out/simple3.ods")?;

    let wb = read_ods("test_out/simple3.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(2, 0).as_i32_or(0), 1);
    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

// basic case, row1 gap row2
// row1 with a repeat of 2
#[test]
fn test_write_gap_repeat() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(2, 0, 1);
    sh.set_row_repeat(2, 2);
    sh.set_value(5, 0, 1);
    wb.push_sheet(sh);

    test_write_ods(&mut wb, "test_out/simple4.ods")?;

    let wb = read_ods("test_out/simple4.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(2, 0).as_i32_or(0), 1);
    assert_eq!(sh.value(5, 0).as_i32_or(0), 1);

    Ok(())
}

#[test]
#[should_panic]
fn test_write_row_overlap() -> () {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(2, 0, 1);
    sh.set_row_repeat(2, 2);
    sh.set_value(3, 0, 1);
    wb.push_sheet(sh);

    match test_write_ods(&mut wb, "test_out/simple5.ods") {
        Ok(_) => {}
        Err(_) => {
            let _ = fs::remove_file("test_out/simple5.ods");
            panic!();
        }
    }
}

#[test]
#[should_panic]
fn test_write_col_overlap() -> () {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(3, 0, 100);
    sh.set_col_repeat(3, 0, 5);
    sh.set_value(3, 4, 101);
    wb.push_sheet(sh);

    match test_write_ods(&mut wb, "test_out/simple6.ods") {
        Ok(_) => {}
        Err(_) => {
            let _ = fs::remove_file("test_out/simple6.ods");
            panic!();
        }
    }
}

#[test]
fn test_write_repeat() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");

    sh.set_value(9, 9, "X");

    sh.set_value(2, 0, 100);
    sh.set_col_repeat(2, 0, 5);

    sh.set_value(3, 0, 100);
    sh.set_col_repeat(3, 0, 20);

    sh.set_value(4, 0, 100);
    sh.set_col_repeat(4, 0, 5);
    sh.set_value(4, 5, 101);

    sh.set_value(5, 1, "V");
    sh.set_col_span(5, 1, 2);
    sh.set_row_span(5, 1, 2);

    sh.set_value(6, 0, 100);
    sh.set_col_repeat(6, 0, 5);
    sh.set_value(6, 5, 101);

    wb.push_sheet(sh);

    test_write_ods(&mut wb, "test_out/col_repeat.ods")?;

    let read = BufReader::new(File::open("test_out/col_repeat.ods")?);
    let wb = OdsOptions::default()
        .use_repeat_for_cells()
        .read_ods(read)?;
    let sh = wb.sheet(0);

    assert_eq!(sh.value(9, 9).as_str_or(""), "X");
    assert_eq!(sh.value(2, 0).as_u32_or(0), 100);
    assert_eq!(sh.col_repeat(2, 0), 5);
    assert_eq!(sh.value(4, 5).as_u32_or(0), 101);
    assert_eq!(sh.value(6, 0).as_u32_or(0), 100);
    assert_eq!(sh.col_repeat(6, 0), 1);
    assert_eq!(sh.value(6, 1).as_u32_or(0), 100);
    assert_eq!(sh.col_repeat(6, 1), 2);
    assert_eq!(sh.value(6, 3).as_u32_or(0), 100);
    assert_eq!(sh.col_repeat(6, 3), 2);
    assert_eq!(sh.value(6, 5).as_u32_or(0), 101);
    assert_eq!(sh.col_repeat(6, 5), 1);

    Ok(())
}
