#![allow(dead_code)]

use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::Path;

use spreadsheet_ods::{OdsError, OdsOptions};

use crate::lib_test::Timing;

mod lib_test;

#[test]
fn test_samples() -> Result<(), OdsError> {
    let mut t1 = run_samples(OdsOptions::default().use_clone_for_cells())?;
    t1.timing.name = "clone".to_string();
    let mut t2 = run_samples(OdsOptions::default().content_only())?;
    t2.timing.name = "content".to_string();
    let mut t3 = run_samples(OdsOptions::default().use_repeat_for_cells())?;
    t3.timing.name = "repeat".to_string();
    let mut t4 = run_samples(OdsOptions::default().ignore_empty_cells())?;
    t4.timing.name = "ignore".to_string();

    print_t(&t1, &t2, &t3, &t4);

    Ok(())
}

fn print_accu(t: &SampleTiming) {
    println!();
    println!("{}", t.timing.name);
    println!();
    println!(
        "| n | sum | 1/10 | median | 9/10 | mean | lin_dev | std_dev | mem-size | size | cells "
    );
    println!("|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|");

    let n = t.timing.n();
    let sum = t.timing.sum();
    let (m0, m5, m9) = t.timing.median();
    let mean = t.timing.mean();
    let lin = t.timing.lin_dev();
    let std = t.timing.std_dev();

    println!(
        "| {} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {:.2}{} | {} | {} | {} |",
        n,
        t.timing.unit.conv(sum),
        t.timing.unit,
        t.timing.unit.conv(m0),
        t.timing.unit,
        t.timing.unit.conv(m5),
        t.timing.unit,
        t.timing.unit.conv(m9),
        t.timing.unit,
        t.timing.unit.conv(mean),
        t.timing.unit,
        t.timing.unit.conv(lin),
        t.timing.unit,
        t.timing.unit.conv(std),
        t.timing.unit,
        t.mem_size.iter().sum::<usize>(),
        t.file_size.iter().sum::<u64>(),
        t.cell_count.iter().sum::<usize>(),
    );
    println!();
}

fn print_t(t0: &SampleTiming, t1: &SampleTiming, t2: &SampleTiming, t3: &SampleTiming) {
    print_accu(t0);
    print_accu(t1);
    print_accu(t2);
    print_accu(t3);

    println!();
    println!("{}", t0.timing.name);
    println!();
    println!("| name | file-size | cells {} | cells {} | cells {} | cells {} | time {} | time {} | time {} | time {} | mem-size {} | mem-size {} | mem-size {} | mem-size {} |",
             t0.timing.name,
             t1.timing.name,
             t2.timing.name,
             t3.timing.name,
             t0.timing.name,
             t1.timing.name,
             t2.timing.name,
             t3.timing.name,
             t0.timing.name,
             t1.timing.name,
             t2.timing.name,
             t3.timing.name,
    );
    for i in 0..t0.timing.samples.len() {
        println!(
            "| {} | {} | {} | {} | {} | {} | {} | {} | {} | {} | {} | {} | {} | {} |",
            t0.name[i],
            t0.file_size[i],
            t0.cell_count[i],
            t1.cell_count[i],
            t2.cell_count[i],
            t3.cell_count[i],
            t0.timing.samples[i],
            t1.timing.samples[i],
            t2.timing.samples[i],
            t3.timing.samples[i],
            t0.mem_size[i],
            t1.mem_size[i],
            t2.mem_size[i],
            t3.mem_size[i],
        );
    }
    // for i in 0..t0.timing.samples.len() {
    //     println!(
    //         "| {} | {} | {} | {} | {} | {} | {} | {} | {} |",
    //         t0.file_size[i],
    //         t0.timing.samples[i] / t0.timing.samples[i],
    //         t1.timing.samples[i] / t0.timing.samples[i],
    //         t2.timing.samples[i] / t0.timing.samples[i],
    //         t3.timing.samples[i] / t0.timing.samples[i],
    //         t0.mem_size[i] / t0.mem_size[i],
    //         t1.mem_size[i] / t0.mem_size[i],
    //         t2.mem_size[i] / t0.mem_size[i],
    //         t3.mem_size[i] / t0.mem_size[i],
    //     );
    // }
}

#[derive(Default)]
struct SampleTiming {
    timing: Timing,
    name: Vec<String>,
    file_size: Vec<u64>,
    cell_count: Vec<usize>,
    mem_size: Vec<usize>,
}

fn run_samples(options: OdsOptions) -> Result<SampleTiming, OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let mut t1 = SampleTiming::default();

    if path.exists() {
        for f in path.read_dir()? {
            let f = f?;

            if f.metadata()?.is_file() {
                if f.file_name().to_string_lossy().ends_with(".ods") {
                    t1.timing.name = f
                        .path()
                        .file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string();

                    let mut buf = Vec::new();
                    File::open(f.path())?.read_to_end(&mut buf)?;

                    let wb = t1.timing.run(|| {
                        let read = BufReader::new(Cursor::new(&buf));
                        options.read_ods(read)
                    })?;

                    let mut cell_count = 0;
                    for sh in wb.iter_sheets() {
                        cell_count += sh.cell_count();
                    }

                    t1.name.push(t1.timing.name.clone());
                    t1.cell_count.push(cell_count);
                    t1.file_size.push(f.metadata()?.len());
                    t1.mem_size.push(loupe::size_of_val(&wb));
                }
            }
        }
    }

    Ok(t1)
}

// #[test]
fn test_sample() -> Result<(), OdsError> {
    let mut t1 = run_sample(OdsOptions::default().use_clone_for_cells())?;
    t1.timing.name = "clone".to_string();
    let mut t2 = run_sample(OdsOptions::default().content_only())?;
    t2.timing.name = "content".to_string();
    let mut t3 = run_sample(OdsOptions::default().use_repeat_for_cells())?;
    t3.timing.name = "repeat".to_string();
    let mut t4 = run_sample(OdsOptions::default().ignore_empty_cells())?;
    t4.timing.name = "ignore".to_string();

    print_t(&t1, &t2, &t3, &t4);

    Ok(())
}

fn run_sample(options: OdsOptions) -> Result<SampleTiming, OdsError> {
    let path = Path::new("C:\\Users\\stommy\\Documents\\StableProjects\\spreadsheet-ods-samples");
    let sample = "1694400557247FBqBqBMJ.ods";

    let f = path.join(sample);

    let mut t1 = SampleTiming::default();
    if f.exists() {
        t1.timing.name = f
            .as_path()
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .to_string();

        let mut buf = Vec::new();
        File::open(&f)?.read_to_end(&mut buf)?;

        let wb = t1.timing.run(|| {
            let read = BufReader::new(Cursor::new(&buf));
            options.read_ods(read)
        })?;

        let mut cell_count = 0usize;
        for sh in wb.iter_sheets() {
            cell_count += sh.cell_count();
        }

        t1.cell_count.push(cell_count);
        t1.name.push(t1.timing.name.clone());
        t1.cell_count.push(cell_count);
        t1.file_size.push(f.metadata()?.len());
        t1.mem_size.push(loupe::size_of_val(&wb));

        println!("cells {}", cell_count);
        println!(
            "sheet {}",
            loupe::size_of_val(&wb.iter_sheets().collect::<Vec<_>>())
        );
        for sh in wb.iter_sheets() {
            println!(
                "col|row header {} | {}",
                sh.col_header_max(),
                sh.row_header_max()
            );
            // println!("extra {}", loupe::size_of_val(sh.extra()));
        }
        println!(
            "font {}",
            loupe::size_of_val(&wb.iter_fonts().collect::<Vec<_>>())
        );
        println!(
            "table-styles {}",
            loupe::size_of_val(&wb.iter_table_styles().collect::<Vec<_>>())
        );
        println!(
            "row-styles {}",
            loupe::size_of_val(&wb.iter_rowstyles().collect::<Vec<_>>())
        );
        println!(
            "col-styles {}",
            loupe::size_of_val(&wb.iter_colstyles().collect::<Vec<_>>())
        );
        println!(
            "cell-styles {}",
            loupe::size_of_val(&wb.iter_cellstyles().collect::<Vec<_>>())
        );
        println!(
            "paragraph-styles {}",
            loupe::size_of_val(&wb.iter_paragraphstyles().collect::<Vec<_>>())
        );
        println!(
            "text-styles {}",
            loupe::size_of_val(&wb.iter_textstyles().collect::<Vec<_>>())
        );
        println!(
            "ruby-styles {}",
            loupe::size_of_val(&wb.iter_rubystyles().collect::<Vec<_>>())
        );
        println!(
            "graphic-styles {}",
            loupe::size_of_val(&wb.iter_graphicstyles().collect::<Vec<_>>())
        );
        println!(
            "boolean-formats {}",
            loupe::size_of_val(&wb.iter_boolean_formats().collect::<Vec<_>>())
        );
        println!(
            "number-formats {}",
            loupe::size_of_val(&wb.iter_number_formats().collect::<Vec<_>>())
        );
        println!(
            "percentage-formats {}",
            loupe::size_of_val(&wb.iter_percentage_formats().collect::<Vec<_>>())
        );
        println!(
            "currency-formats {}",
            loupe::size_of_val(&wb.iter_currency_formats().collect::<Vec<_>>())
        );
        println!(
            "text-formats {}",
            loupe::size_of_val(&wb.iter_text_formats().collect::<Vec<_>>())
        );
        println!(
            "datetime-formats {}",
            loupe::size_of_val(&wb.iter_datetime_formats().collect::<Vec<_>>())
        );
        println!(
            "timeduration-formats {}",
            loupe::size_of_val(&wb.iter_timeduration_formats().collect::<Vec<_>>())
        );
        println!(
            "number-formats {}",
            loupe::size_of_val(&wb.iter_number_formats().collect::<Vec<_>>())
        );
        println!(
            "page-styles {}",
            loupe::size_of_val(&wb.iter_pagestyles().collect::<Vec<_>>())
        );
        println!(
            "masterpages {}",
            loupe::size_of_val(&wb.iter_masterpages().collect::<Vec<_>>())
        );
        println!(
            "validations {}",
            loupe::size_of_val(&wb.iter_validations().collect::<Vec<_>>())
        );
        println!("config {}", loupe::size_of_val(&wb.config()));
        println!(
            "manifest {}",
            loupe::size_of_val(&wb.iter_manifest().collect::<Vec<_>>())
        );
        println!("metadata {}", loupe::size_of_val(&wb.metadata()));
        // println!("extra {}", loupe::size_of_val(&wb.extra()));
        // println!(
        //     "workbook-config {}",
        //     loupe::size_of_val(&wb.workbook_config())
        // );
    }

    Ok(t1)
}
