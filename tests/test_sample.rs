use std::path::Path;
use std::time::{Duration, Instant};

use spreadsheet_ods::{read_ods, OdsError};

// fn stat_sheet(sh: &Sheet) -> (u32, u32, u32, u32, u32, u32) {
//     let mut n_empty = 0;
//     let mut n_value = 0;
//     let mut n_formula = 0;
//     let mut n_style = 0;
//     let mut n_span = 0;
//
//     let (rows, cols) = sh.used_grid_size();
//
//     for r in 0..rows {
//         for c in 0..cols {
//             match sh.value(r, c) {
//                 Value::Empty => n_empty += 1,
//                 _ => n_value += 1,
//             };
//             if sh.formula(r, c).is_some() {
//                 n_formula += 1;
//             }
//             if sh.cellstyle(r, c).is_some() {
//                 n_style += 1;
//             }
//             if sh.col_span(r, c) > 1 || sh.row_span(r, c) > 1 {
//                 n_span += 1;
//             }
//         }
//     }
//
//     (rows * cols, n_empty, n_value, n_formula, n_style, n_span)
// }

fn timingr<E, R>(
    name: &str,
    mut fun: impl FnMut() -> Result<R, E>,
    dur: &mut Duration,
) -> Result<R, E> {
    let now = Instant::now();
    let result = fun()?;
    *dur = now.elapsed();
    println!("{} {:?}", name, dur);
    Ok(result)
}

#[test]
fn test_samples() -> Result<(), OdsError> {
    let path = Path::new("..\\spreadsheet-ods-samples\\");

    if path.exists() {
        let mut dur = Duration::new(0, 0);
        let mut count = 0;

        for f in path.read_dir()? {
            let f = f?;

            if f.metadata()?.is_file() {
                let name = f.file_name();
                if name.to_string_lossy().ends_with(".ods") {
                    let ff = path.join(name);

                    println!();
                    println!("{:?}", ff);

                    let mut xdur = Duration::new(0, 0);
                    let _wb = timingr("read_ods", || read_ods(&ff), &mut xdur)?;

                    dur += xdur;
                    count += 1;
                }
            }
        }

        println!();
        println!(
            "avg {}ms count {}",
            dur.as_nanos() as f64 / count as f64 / 1e6,
            count
        );
    }

    Ok(())
}

#[test]
fn test_sample() -> Result<(), OdsError> {
    let sample = "martinique_logementssociauxvides_1erjanvier2020_sup_6.ods";
    let path = Path::new("..\\spreadsheet-ods-samples\\");

    let mut dur = Duration::new(0, 0);

    let ff = path.join(sample);

    println!();
    println!("{:?}", ff);

    let mut xdur = Duration::new(0, 0);
    let wb = timingr("read_ods", || read_ods(&ff), &mut xdur)?;
    dur += xdur;

    for v in wb.iter_rubystyles() {
        dbg!(v);
    }

    println!();
    println!("dur {}ms", dur.as_nanos() as f64 / 1e6,);

    Ok(())
}
