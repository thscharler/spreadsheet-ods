use std::path::Path;
use std::time::{Duration, Instant};

use spreadsheet_ods::{read_ods, OdsError};

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
                    let _wb = timingr(
                        format!("read {:?} {}", ff, f.metadata()?.len()).as_str(),
                        || read_ods(&ff),
                        &mut xdur,
                    )?;

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
    let path = Path::new("..\\spreadsheet-ods-samples\\");

    let sample = "108_BasicChange.ods";

    let mut dur = Duration::new(0, 0);

    let ff = path.join(sample);

    println!();
    println!("{:?}", ff);

    let mut xdur = Duration::new(0, 0);
    let wb = timingr(
        format!("read {:?} {}", ff, ff.metadata()?.len()).as_str(),
        || read_ods(&ff),
        &mut xdur,
    )?;
    dur += xdur;

    println!();
    println!("dur {}ms", dur.as_nanos() as f64 / 1e6,);

    Ok(())
}
