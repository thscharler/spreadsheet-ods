//!
//! Defines types for cell references.
//!

use spreadsheet_ods_cellref::error::CellRefError;
use spreadsheet_ods_cellref::refs_format::fmt_cell_range_list;
use spreadsheet_ods_cellref::refs_parser::{parse_cell_range_list, parse_cell_ref, Span};
use std::fmt;
use std::fmt::{Display, Formatter, Write};

pub use spreadsheet_ods_cellref::refs::{CellRange, CellRef, ColRange, RowRange};

/// Parse a cell reference.
pub fn parse_cellref(buf: &str, pos: &mut usize) -> Result<CellRef, CellRefError> {
    let (_rest, (cell_ref, _tok)) = parse_cell_ref(Span::new(&buf[*pos..]))?;
    Ok(cell_ref)
}

/// Parse a list of range refs
pub fn parse_cellranges(
    buf: &str,
    pos: &mut usize,
) -> Result<Option<Vec<CellRange>>, CellRefError> {
    let (_rest, vec) = parse_cell_range_list(Span::new(&buf[*pos..]))?;
    Ok(vec)
}

/// Returns a list of ranges as string.
pub fn cellranges_string(vec: &[CellRange]) -> String {
    let mut buf = String::new();
    let _ = write!(buf, "{}", Fmt(|f| fmt_cell_range_list(f, vec)));
    buf
}

struct Fmt<F>(F)
where
    for<'a> F: Fn(&mut Formatter<'a>) -> fmt::Result;

impl<F> Display for Fmt<F>
where
    for<'a> F: Fn(&mut Formatter<'a>) -> fmt::Result,
{
    /// Calls f with the given Formatter.
    fn fmt<'a>(&self, f: &mut Formatter<'a>) -> fmt::Result {
        (self.0)(f)
    }
}
