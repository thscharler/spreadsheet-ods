use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::Path;
use std::time::{Duration, Instant};

use spreadsheet_ods::{read_ods, OdsError, OdsOptions};

fn timing_run<E, R>(
    name: &str,
    mut fun: impl FnMut() -> Result<R, E>,
    repeat: u32,
) -> Result<(Duration, R), E> {
    let mut result = None;
    let now = Instant::now();
    for _ in 0..repeat {
        result = Some(fun()?);
    }
    let dur = now.elapsed();
    println!("{} {:?}|{}", name, dur / repeat, repeat);
    Ok((dur, result.unwrap()))
}

#[test]
fn test_samples() -> Result<(), OdsError> {
    // let path = Path::new("..\\spreadsheet-ods-samples\\");
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");

    if path.exists() {
        let mut total = Duration::default();
        let mut count = 0;

        for f in path.read_dir()? {
            let f = f?;

            if f.metadata()?.is_file() {
                if f.file_name().to_string_lossy().ends_with(".ods") {
                    println!();
                    println!("{:?} {}", f.path(), f.metadata()?.len());

                    let mut buf = Vec::new();
                    File::open(f.path())?.read_to_end(&mut buf)?;

                    let (dur, _) = timing_run(
                        "read",
                        || {
                            let read = BufReader::new(Cursor::new(&buf));
                            OdsOptions::default().read_ods(read)
                        },
                        1,
                    )?;

                    total += dur;
                    count += 1;
                }
            }
        }

        //

        let mut total2 = Duration::default();
        let mut count2 = 0;

        for f in path.read_dir()? {
            let f = f?;

            if f.metadata()?.is_file() {
                if f.file_name().to_string_lossy().ends_with(".ods") {
                    println!();
                    println!("{:?} {}", f.path(), f.metadata()?.len());

                    let mut buf = Vec::new();
                    File::open(f.path())?.read_to_end(&mut buf)?;

                    let (dur, _) = timing_run(
                        "read",
                        || {
                            let read = BufReader::new(Cursor::new(&buf));
                            OdsOptions::default().content_only().read_ods(read)
                        },
                        1,
                    )?;

                    total2 += dur;
                    count2 += 1;
                }
            }
        }

        println!("{:?} {} avg {:?}", total, count, total / count);
        println!("{:?} {} avg {:?}", total2, count2, total2 / count2);
    }

    Ok(())
}

#[test]
fn test_sample() -> Result<(), OdsError> {
    // let path = Path::new("..\\spreadsheet-ods-samples\\");
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let sample = "businesstrip11201.ods";

    let f = path.join(sample);

    println!();
    println!("{:?} {}", f.as_path(), f.metadata()?.len());

    let mut buf = Vec::new();
    File::open(&f)?.read_to_end(&mut buf)?;

    timing_run(
        "read",
        || {
            let read = BufReader::new(Cursor::new(&buf));
            OdsOptions::default().read_ods(read)
        },
        1,
    )?;
    timing_run(
        "reac",
        || {
            let read = BufReader::new(Cursor::new(&buf));
            OdsOptions::default().content_only().read_ods(read)
        },
        1,
    )?;

    Ok(())
}
