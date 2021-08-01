//!
//! For now defines functions to create cell references for formulas.
//!

use crate::refs::{CellRange, CellRef};

// TODO: more formula stuff. parsing?

/// Creates a cell-reference for use in formulas.
pub fn fcellref(row: u32, col: u32) -> String {
    CellRef::local(row, col).to_formula()
}

/// Creates a cell-reference for use in formulas.
/// Creates an absolute row reference.
pub fn fcellrefr(row: u32, col: u32) -> String {
    CellRef::local(row, col).absolute_row().to_formula()
}

/// Creates a cell-reference for use in formulas.
/// Creates an absolute col reference.
pub fn fcellrefc(row: u32, col: u32) -> String {
    CellRef::local(row, col).absolute_col().to_formula()
}

/// Creates a cell-reference for use in formulas.
/// Creates an absolute reference.
pub fn fcellrefa(row: u32, col: u32) -> String {
    CellRef::local(row, col).absolute().to_formula()
}

/// Creates a cell-reference for use in formulas.
pub fn fcellref_table<S: Into<String>>(table: S, row: u32, col: u32) -> String {
    CellRef::remote(table, row, col).to_formula()
}

/// Creates a cell-reference for use in formulas.
/// Creates an absolute row reference.
pub fn fcellrefr_table<S: Into<String>>(table: S, row: u32, col: u32) -> String {
    CellRef::remote(table, row, col).absolute_row().to_formula()
}

/// Creates a cell-reference for use in formulas.
/// Creates an absolute col reference.
pub fn fcellrefc_table<S: Into<String>>(table: S, row: u32, col: u32) -> String {
    CellRef::remote(table, row, col).absolute_col().to_formula()
}

/// Creates a cell-reference for use in formulas.
/// Creates an absolute reference.
pub fn fcellrefa_table<S: Into<String>>(table: S, row: u32, col: u32) -> String {
    CellRef::remote(table, row, col).absolute().to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangeref(row: u32, col: u32, row_to: u32, col_to: u32) -> String {
    CellRange::local(row, col, row_to, col_to).to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangerefr(row: u32, col: u32, row_to: u32, col_to: u32) -> String {
    CellRange::local(row, col, row_to, col_to)
        .absolute_rows()
        .to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangerefc(row: u32, col: u32, row_to: u32, col_to: u32) -> String {
    CellRange::local(row, col, row_to, col_to)
        .absolute_cols()
        .to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangerefa(row: u32, col: u32, row_to: u32, col_to: u32) -> String {
    CellRange::local(row, col, row_to, col_to)
        .absolute()
        .to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangeref_table<S: Into<String>>(
    table: S,
    row: u32,
    col: u32,
    row_to: u32,
    col_to: u32,
) -> String {
    CellRange::remote(table, row, col, row_to, col_to).to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangerefr_table<S: Into<String>>(
    table: S,
    row: u32,
    col: u32,
    row_to: u32,
    col_to: u32,
) -> String {
    CellRange::remote(table, row, col, row_to, col_to)
        .absolute_rows()
        .to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangerefc_table<S: Into<String>>(
    table: S,
    row: u32,
    col: u32,
    row_to: u32,
    col_to: u32,
) -> String {
    CellRange::remote(table, row, col, row_to, col_to)
        .absolute_cols()
        .to_formula()
}

/// Creates a cellrange-reference for use in formulas.
pub fn frangerefa_table<S: Into<String>>(
    table: S,
    row: u32,
    col: u32,
    row_to: u32,
    col_to: u32,
) -> String {
    CellRange::remote(table, row, col, row_to, col_to)
        .absolute()
        .to_formula()
}
