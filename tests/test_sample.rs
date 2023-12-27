use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::Path;

use crate::lib_test::Timing;
use spreadsheet_ods::{OdsError, OdsOptions};

mod lib_test;

#[test]
fn test_samples() -> Result<(), OdsError> {
    let t1 = run_samples(OdsOptions::default().use_clone_for_cells())?.name("clone");
    let t2 = run_samples(OdsOptions::default().content_only())?.name("content");
    let t3 = run_samples(OdsOptions::default().use_repeat_for_cells())?.name("repeat");
    let t4 = run_samples(OdsOptions::default().ignore_empty_cells())?.name("ignore");

    println!("{}", t1);
    println!("{}", t2);
    println!("{}", t3);
    println!("{}", t4);

    Ok(())
}

fn run_samples(options: OdsOptions) -> Result<Timing, OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let mut t1 = Timing::default();

    if path.exists() {
        for f in path.read_dir()? {
            let f = f?;

            if f.metadata()?.is_file() {
                if f.file_name().to_string_lossy().ends_with(".ods") {
                    println!();
                    println!("{:?} {}", f.path(), f.metadata()?.len());

                    let mut buf = Vec::new();
                    File::open(f.path())?.read_to_end(&mut buf)?;

                    t1.run(|| {
                        let read = BufReader::new(Cursor::new(&buf));
                        options.read_ods(read)
                    })?;
                }
            }
        }
    }

    Ok(t1)
}

#[test]
fn test_sample() -> Result<(), OdsError> {
    run_sample(OdsOptions::default().use_clone_for_cells())
}

#[test]
fn test_sample_content() -> Result<(), OdsError> {
    run_sample(OdsOptions::default().content_only())
}

#[test]
fn test_sample_repeat() -> Result<(), OdsError> {
    run_sample(OdsOptions::default().use_repeat_for_cells())
}

#[test]
fn test_sample_ignore() -> Result<(), OdsError> {
    run_sample(OdsOptions::default().ignore_empty_cells())
}

fn run_sample(options: OdsOptions) -> Result<(), OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let sample = "spanimals12-supptabs.ods";

    let f = path.join(sample);

    if f.exists() {
        let mut t1 = Timing::default();

        println!();
        println!("{:?} {}", f.as_path(), f.metadata()?.len());

        let mut buf = Vec::new();
        File::open(&f)?.read_to_end(&mut buf)?;

        let wb = t1.run(|| {
            let read = BufReader::new(Cursor::new(&buf));
            options.read_ods(read)
        })?;

        let mut cell_count = 0usize;
        for sh in wb.iter_sheets() {
            cell_count += sh.cell_count();
        }

        println!("cell_count {}", cell_count);
        println!("{}", t1);
    }

    Ok(())
}
