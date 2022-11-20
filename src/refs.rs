//!
//! Defines types for cell references.
//!

use crate::refs_impl::error::OFCode;
use crate::refs_impl::{check_eof, parser, Span};
use crate::OdsError;
pub use spreadsheet_ods_cellref::{CellRange, CellRef, ColRange, RowRange};
use std::fmt;
use std::fmt::{Display, Formatter, Write};

/// Parse a cell reference.
pub fn parse_cellref(buf: &str, _pos: &mut usize) -> Result<CellRef, OdsError> {
    let rest = Span::new(buf);

    let (rest, tok) = crate::refs_impl::parser::parse_cell_ref(rest)?;

    check_eof(rest, OFCode::OFCCellRef)?;

    Ok(CellRef::new_all(
        tok.iri.map(|v| v.iri),
        tok.table.map(|v| v.name),
        tok.row.abs,
        tok.row.row,
        tok.col.abs,
        tok.col.col,
    ))
}

/// Parse a list of range refs
pub fn parse_cellranges(buf: &str, _pos: &mut usize) -> Result<Option<Vec<CellRange>>, OdsError> {
    let rest = Span::new(buf);

    let (rest, ranges) = parser::parse_cell_range_list(rest)?;

    check_eof(rest, OFCode::OFCCellRef)?;

    let ranges = ranges.map(|o| {
        o.into_iter()
            .map(|tok| {
                CellRange::new_all(
                    tok.iri.map(|v| v.iri),
                    tok.table.map(|v| v.name),
                    tok.row.abs,
                    tok.row.row,
                    tok.col.abs,
                    tok.col.col,
                    tok.to_table.map(|v| v.name),
                    tok.to_row.abs,
                    tok.to_row.row,
                    tok.to_col.abs,
                    tok.to_col.col,
                )
            })
            .collect()
    });

    Ok(ranges)
}

/// Returns a list of ranges as string.
pub fn cellranges_string(vec: &[CellRange]) -> String {
    let mut buf = String::new();

    for range in vec {
        let _err = write!(buf, "{}", range);
    }

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
