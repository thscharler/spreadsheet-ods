use std::mem::{size_of, size_of_val};
use std::time::Instant;

use spreadsheet_ods::defaultstyles::create_default_styles;
use spreadsheet_ods::{
    ucell, write_ods_buf_uncompressed, ColHeader, OdsError, RowHeader, SCell, Sheet, WorkBook,
};

pub fn timingr<E, R>(
    name: &str,
    divider: u64,
    mut fun: impl FnMut() -> Result<R, E>,
) -> Result<R, E> {
    let now = Instant::now();
    println!("{}", name);
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
        let mut wb = WorkBook::new();
        create_default_styles(&mut wb);
        let mut sh = Sheet::new();

        for r in 0..rows {
            for c in 0..cols {
                sh.set_value(r, c, 1);
            }
        }

        wb.push_sheet(sh);

        Ok(wb)
    }
}

fn write_wb<'a>(wb: &'a mut WorkBook) -> impl FnMut() -> Result<(), OdsError> + 'a {
    move || {
        let buf = write_ods_buf_uncompressed(wb, Vec::new())?;
        println!("len {}", buf.len());
        Ok(())
    }
}

#[test]
fn test_b0() -> Result<(), OdsError> {
    const ROWS: u32 = 10000;
    const COLS: u32 = 40;
    const CELLS: u64 = ROWS as u64 * COLS as u64;
    println!("{}", ROWS * COLS);
    let mut wb = timingr(
        "create_wb",
        ROWS as u64 * COLS as u64,
        create_wb(ROWS, COLS),
    )?;
    timingr("write_wb", CELLS, write_wb(&mut wb))?;

    let size = size_of::<(ucell, ucell)>() * CELLS as usize
        + size_of::<SCell>() * CELLS as usize
        + (size_of::<ColHeader>() + size_of::<ucell>()) * COLS as usize
        + (size_of::<RowHeader>() + size_of::<ucell>()) * ROWS as usize;
    println!("ca mem {}", size);

    Ok(())
}
