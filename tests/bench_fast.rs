mod lib_test;

use crate::lib_test::Timing;
use icu_locid::locale;
use spreadsheet_ods::sheet::Visibility;
use spreadsheet_ods::{
    write_ods_buf, write_ods_buf_uncompressed, Length, OdsError, Sheet, WorkBook,
};

#[allow(dead_code)]
fn create_wb(rows: u32, cols: u32) -> impl FnMut() -> Result<WorkBook, OdsError> {
    move || {
        let mut wb = WorkBook::new_empty();
        wb.locale_settings(locale!("en_US"));
        let mut sh = Sheet::new("1");

        for r in 0..rows {
            if r % 2 == 0 {
                for c in 0..cols {
                    sh.set_value(r, c, "1234");
                }
            } else {
                for c in 0..cols {
                    sh.set_value(r, c, 1u32);
                }
            }
            if r % 2 == 0 {
                for c in 0..cols {
                    sh.set_cellstyle(r, c, &"s0".into());
                }
            }
            if r % 10 == 0 {
                for c in 0..cols {
                    sh.set_formula(r, c, "of:=1+1");
                }
            }
            if r % 50 == 0 {
                for c in 0..cols {
                    sh.set_validation(r, c, &"v0".into());
                }
            }
        }

        wb.push_sheet(sh);

        Ok(wb)
    }
}

#[allow(dead_code)]
fn write_wb<'a>(wb: &'a mut WorkBook) -> impl FnMut() -> Result<(), OdsError> + 'a {
    move || {
        let buf = write_ods_buf_uncompressed(wb, Vec::new())?;
        println!("len {}", buf.len());
        Ok(())
    }
}

#[allow(dead_code)]
struct DummyColHeader {
    style: Option<String>,
    cellstyle: Option<String>,
    visible: Visibility,
    width: Length,
}

#[allow(dead_code)]
struct DummyRowHeader {
    style: Option<String>,
    cellstyle: Option<String>,
    visible: Visibility,
    repeat: u32,
    height: Length,
}

#[test]
#[allow(dead_code)]
fn test_b0() -> Result<(), OdsError> {
    const ROWS: u32 = 100;
    const COLS: u32 = 400;

    let mut t1 = Timing::default()
        .name("create_wb")
        .skip(2)
        .runs(30)
        .divider(ROWS as u64 * COLS as u64);
    let mut t2 = t1.clone().name("write_wb");

    let mut wb = t1.run(create_wb(ROWS, COLS))?;
    let _ = t2.run(|| write_ods_buf(&mut wb, Vec::new()));

    println!("{}", t1);
    println!("{}", t2);

    Ok(())
}
