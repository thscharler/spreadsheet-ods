use icu_locid::locale;
use std::time::Instant;

use spreadsheet_ods::{write_ods_buf, Length, OdsError, Sheet, Visibility, WorkBook};

pub fn timingr<E, R>(
    name: &str,
    divider: u64,
    mut fun: impl FnMut() -> Result<R, E>,
) -> Result<R, E> {
    let now = Instant::now();
    // println!("{}", name);
    let result = fun()?;
    let elapsed = now.elapsed();
    println!(
        "{} {:?} {:?}ns/{}",
        name,
        elapsed,
        elapsed.as_nanos() / divider as u128,
        divider
    );
    Ok(result)
}

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

fn write_wb(wb: &mut WorkBook) -> impl FnMut() -> Result<(), OdsError> + '_ {
    move || {
        let buf = write_ods_buf(wb, Vec::new())?;
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
fn test_b0() -> Result<(), OdsError> {
    const ROWS: u32 = 100;
    const COLS: u32 = 40;
    const CELLS: u64 = ROWS as u64 * COLS as u64;

    // println!("sizes {}", size_of::<Value>());

    // println!("{}", ROWS * COLS);
    let mut wb = timingr(
        "create_wb",
        ROWS as u64 * COLS as u64,
        create_wb(ROWS, COLS),
    )?;
    timingr("write_wb", CELLS, write_wb(&mut wb))?;

    Ok(())
}
