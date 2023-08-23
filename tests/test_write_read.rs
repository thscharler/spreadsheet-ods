use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use spreadsheet_ods::{
    read_ods, read_ods_buf, write_ods, write_ods_buf, write_ods_to, OdsError, Sheet, SplitMode,
    ValueType, WorkBook,
};
use std::time::Instant;

// #[test]
// fn test_read2() {
//     let wb = read_ods("tests/Kunden.ods").unwrap();
//     assert_eq!(wb.num_sheets(), 2);
// }

#[test]
fn test_write_read() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();
    let mut sh = Sheet::new("1");

    sh.set_value(0, 0, "A");

    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/test_0.ods")?;

    let wi = read_ods("test_out/test_0.ods")?;
    let si = wi.sheet(0);

    assert_eq!(si.value(0, 0).as_str_or(""), "A");

    Ok(())
}

#[test]
fn read_text() -> Result<(), OdsError> {
    let wb = read_ods("tests/text.ods")?;
    let sh = wb.sheet(0);

    let v = sh.value(0, 0);

    assert_eq!(v.value_type(), ValueType::TextXml);

    Ok(())
}

pub fn timingr<E, R>(name: &str, mut fun: impl FnMut() -> Result<R, E>) -> Result<R, E> {
    let now = Instant::now();
    let result = fun()?;
    println!("{} {:?}", name, now.elapsed());
    Ok(result)
}

pub fn timingn<E>(name: &str, mut fun: impl FnMut()) -> Result<(), E> {
    let now = Instant::now();
    fun();
    println!("{} {:?}", name, now.elapsed());
    Ok(())
}

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

#[test]
fn test_write_read_write_read() -> Result<(), OdsError> {
    let path = Path::new("tests/rw.ods");
    let temp = Path::new("test_out/rw.ods");

    std::fs::copy(path, temp)?;

    let mut ods = read_ods(temp)?;
    write_ods(&mut ods, temp)?;
    let _ods = read_ods(temp)?;

    Ok(())
}

#[test]
fn test_write_repeat_overlapped() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();
    let mut sh = Sheet::new("1");

    sh.set_value(0, 0, "A");
    sh.set_row_repeat(0, 3);
    sh.set_value(1, 0, "X");
    sh.set_value(2, 0, "X");
    sh.set_value(3, 0, "B");

    wb.push_sheet(sh);

    let path = Path::new("test_out/overlap.ods");
    write_ods(&mut wb, path)?;

    let _ods = read_ods(path)?;

    Ok(())
}

#[test]
fn test_write_buf() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();
    let mut sh = Sheet::new("1");

    sh.set_value(0, 0, "A");
    wb.push_sheet(sh);

    let p = Path::new("test_out/bufnot.ods");
    write_ods(&mut wb, p)?;
    let len = p.to_path_buf().metadata()?.len() as usize;

    let v = Vec::new();
    let v = write_ods_buf(&mut wb, v)?;

    assert_eq!(v.len(), len);

    let mut ff = File::create("test_out/bufbuf.ods")?;
    ff.write_all(&v)?;

    Ok(())
}

#[test]
fn test_read_buf() -> Result<(), OdsError> {
    let mut buf = Vec::new();
    let mut f = File::open("tests/orders.ods")?;
    f.read_to_end(&mut buf)?;

    let mut wb = read_ods_buf(&buf)?;

    wb.config_mut().has_sheet_tabs = false;

    let cc = wb.sheet_mut(0).config_mut();
    cc.show_grid = true;
    cc.vert_split_pos = 2;
    cc.vert_split_mode = SplitMode::Heading;

    write_ods(&mut wb, "test_out/orders.ods")?;
    Ok(())
}

#[test]
fn test_write_write() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();
    let mut sh = Sheet::new("1");

    sh.set_value(0, 0, "A");
    wb.push_sheet(sh);

    let v = Cursor::new(Vec::new());
    write_ods_to(&mut wb, v)?;

    Ok(())
}
