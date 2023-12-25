use icu_locid::locale;
use std::hint::black_box;
use std::time::Instant;

use spreadsheet_ods::sheet::Visibility;
use spreadsheet_ods::{
    write_ods_buf, write_ods_buf_uncompressed, Length, OdsError, Sheet, WorkBook,
};

#[derive(Clone, Copy, Debug)]
pub struct Timing {
    pub skip: usize,
    pub runs: usize,
    pub divider: u64,
}

pub fn timingr<E, R>(
    name: &str,
    timing: Timing,
    mut fun: impl FnMut() -> Result<R, E>,
) -> Result<R, E> {
    assert!(timing.runs > 0);
    assert!(timing.divider > 0);

    let mut bench = move || {
        let now = Instant::now();
        let result = fun();
        (now.elapsed(), result)
    };

    let mut elapsed_vec = Vec::with_capacity(timing.runs);
    let mut n = 0;
    let result = loop {
        let (elapsed, result) = black_box(bench());
        elapsed_vec.push(elapsed);
        n += 1;
        if n >= timing.runs + timing.skip {
            break result;
        }
    };

    let elapsed_vec = elapsed_vec
        .iter()
        .skip(timing.skip)
        .map(|v| v.as_nanos() as f64 / timing.divider as f64)
        .collect::<Vec<_>>();

    let mean = elapsed_vec.iter().sum::<f64>() / timing.runs as f64;

    let lin_sum = elapsed_vec.iter().map(|v| (*v - mean).abs()).sum::<f64>();
    let lin_dev = lin_sum / timing.runs as f64;

    let std_sum = elapsed_vec
        .iter()
        .map(|v| (*v - mean) * (*v - mean))
        .sum::<f64>();
    let std_dev = (std_sum / timing.runs as f64).sqrt();

    println!();
    println!("{}", name);
    println!();
    println!("| mean | lin_dev | std_dev |");
    println!("|:---|:---|:---|");
    println!("| {:.2} | {:.2} | {:.2} |", mean, lin_dev, std_dev);
    println!();

    for i in 0..elapsed_vec.len() {
        print!("| {} ", i);
    }
    println!("|");
    for _ in 0..elapsed_vec.len() {
        print!("|:---");
    }
    println!("|");
    for e in &elapsed_vec {
        print!("| {:.2} ", e);
    }
    println!("|");
    for e in &elapsed_vec {
        print!("| {:.2} ", e - mean);
    }
    println!("|");

    result
}

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

// #[test]
#[allow(dead_code)]
fn test_b0() -> Result<(), OdsError> {
    const ROWS: u32 = 100;
    const COLS: u32 = 400;

    let t = Timing {
        skip: 2,
        runs: 30,
        divider: ROWS as u64 * COLS as u64,
    };

    let mut wb = timingr("create_wb", t, create_wb(ROWS, COLS))?;
    let _ = timingr("write_wb", t, || write_ods_buf(&mut wb, Vec::new()));

    Ok(())
}
