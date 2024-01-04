#![allow(dead_code)]

use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::Path;

use spreadsheet_ods::{OdsError, OdsOptions};

use crate::lib_test::Timing;

mod lib_test;

fn print_accu(t: &Timing<Sample>) {
    println!();
    println!("{}", t.name);
    println!();
    println!(
        "| n | sum | 1/10 | median | 9/10 | mean | lin_dev | std_dev | mem-size | size | cells "
    );
    println!("|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|");

    let n = t.n();
    let sum = t.sum();
    let (m0, m5, m9) = t.median();
    let mean = t.mean();
    let lin = t.lin_dev();
    let std = t.std_dev();

    println!(
        "| {} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {} | {} | {} |",
        n,
        t.unit.conv(sum),
        t.unit,
        t.unit.conv(m0),
        t.unit,
        t.unit.conv(m5),
        t.unit,
        t.unit.conv(m9),
        t.unit,
        t.unit.conv(mean),
        t.unit,
        t.unit.conv(lin),
        t.unit,
        t.unit.conv(std),
        t.unit,
        t.extra.iter().fold(0, |s, v| s + v.mem_size),
        t.extra.iter().fold(0, |s, v| s+ v.file_size),
        t.extra.iter().fold(0, |s, v| s+ v.cell_count),
    );
    println!();
}

fn print_t(t0: &Timing<Sample>) {
    print_accu(t0);

    println!();
    println!("{}", t0.name);
    println!();
    println!("| cat | name | file-size | time | cells | mem-size | sheet | colh | rowh | meta |");
    for i in 0..t0.samples.len() {
        let extra = t0.extra.get(i).expect("b");
        println!(
            "| {} | {} | {} | {} | {} | {} | {} | {} | {} | {} |",
            extra.category,
            extra.name,
            extra.file_size,
            t0.samples[i],
            extra.cell_count,
            extra.mem_size,
            extra.sheet_size,
            extra.col_header,
            extra.row_header,
            extra.metadata,
        );
    }
}

#[derive(Default)]
struct Sample {
    category: String,
    name: String,
    file_size: u64,
    cell_count: usize,
    mem_size: usize,
    sheet_size: usize,
    col_header: usize,
    row_header: usize,
    metadata: usize,
}

#[test]
fn test_samples() -> Result<(), OdsError> {
    let mut t = Timing::default();

    run_samples(&mut t, "clone", OdsOptions::default().use_clone_for_cells())?;
    // run_samples(&mut t, "content", OdsOptions::default().content_only())?;
    run_samples(
        &mut t,
        "repeat",
        OdsOptions::default().use_repeat_for_cells(),
    )?;
    run_samples(&mut t, "ignore", OdsOptions::default().ignore_empty_cells())?;

    print_t(&t);

    Ok(())
}

fn run_samples(t1: &mut Timing<Sample>, cat: &str, options: OdsOptions) -> Result<(), OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    if path.exists() {
        for f in path.read_dir()? {
            let f = f?;

            if f.metadata()?.is_file() {
                if f.file_name().to_string_lossy().ends_with(".ods") {
                    t1.name = f
                        .path()
                        .file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string();

                    let mut buf = Vec::new();
                    File::open(f.path())?.read_to_end(&mut buf)?;

                    let wb = t1.run(|| {
                        let read = BufReader::new(Cursor::new(&buf));
                        options.read_ods(read)
                    })?;

                    let mut cell_count = 0;
                    for sh in wb.iter_sheets() {
                        cell_count += sh.cell_count();
                    }

                    t1.extra.push(Sample {
                        category: cat.to_string(),
                        name: t1.name.clone(),
                        file_size: f.metadata()?.len(),
                        cell_count,
                        mem_size: loupe::size_of_val(&wb),
                        ..Sample::default()
                    });
                }
            }
        }
    }

    Ok(())
}

#[test]
fn test_sample() -> Result<(), OdsError> {
    let mut t = Timing::default();
    run_sample(&mut t, "clone", OdsOptions::default().use_clone_for_cells())?;
    // run_sample(&mut t, "content", OdsOptions::default().content_only())?;
    // run_sample(
    //     &mut t,
    //     "repeat",
    //     OdsOptions::default().use_repeat_for_cells(),
    // )?;
    // run_sample(&mut t, "ignore", OdsOptions::default().ignore_empty_cells())?;

    print_t(&t);

    Ok(())
}

fn run_sample(t1: &mut Timing<Sample>, cat: &str, options: OdsOptions) -> Result<(), OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let sample = "2d0d3aca0b2ddd244ad34f2b11f5625cd2835141ca98ec025a54c2f2d10118.ods";

    let f = path.join(sample);
    if f.exists() {
        t1.name = f
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();

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

        t1.extra.push(Sample {
            category: cat.to_string(),
            name: t1.name.clone(),
            file_size: f.metadata()?.len(),
            cell_count,
            mem_size: loupe::size_of_val(&wb),
            sheet_size: loupe::size_of_val(&wb.iter_sheets().collect::<Vec<_>>()),
            col_header: wb.iter_sheets().map(|v| v._col_header_len()).sum(),
            row_header: wb.iter_sheets().map(|v| v._row_header_len()).sum(),
            metadata: loupe::size_of_val(&wb.iter_fonts().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_table_styles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_rowstyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_colstyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_cellstyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_paragraphstyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_textstyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_rubystyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_graphicstyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_boolean_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_number_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_percentage_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_currency_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_text_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_datetime_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_timeduration_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_number_formats().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_pagestyles().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_masterpages().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.iter_validations().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.config())
                + loupe::size_of_val(&wb.iter_manifest().collect::<Vec<_>>())
                + loupe::size_of_val(&wb.metadata()),
        });
    }

    Ok(())
}

#[test]
fn run_dump() -> Result<(), OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let sample = "2d0d3aca0b2ddd244ad34f2b11f5625cd2835141ca98ec025a54c2f2d10118.ods";

    let mut buf = Vec::new();

    let f = path.join(sample);
    File::open(&f)?.read_to_end(&mut buf)?;

    let read = BufReader::new(Cursor::new(&buf));
    let wb = OdsOptions::default().read_ods(read)?;

    let sh = wb.sheet(0);

    dbg!(sh);

    Ok(())
}
