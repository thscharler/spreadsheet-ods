use crate::{CellRange, CellRef, ucell};

/// Returns the spreadsheet column name.
pub fn push_colname(buf: &mut String, mut col: ucell) {
    let mut i = 0;
    let mut dbuf = [0u8; 7];

    col += 1;
    while col > 0 {
        dbuf[i] = (col % 26) as u8;
        if dbuf[i] == 0 {
            dbuf[i] = 25;
            col = col / 26 - 1;
        } else {
            dbuf[i] -= 1;
            col /= 26;
        }

        i += 1;
    }

    // reverse order
    let mut j = i;
    while j > 0 {
        buf.push((b'A' + dbuf[j - 1]) as char);
        j -= 1;
    }
}

/// Returns the spreadsheet row name
pub fn push_rowname(buf: &mut String, mut row: ucell) {
    let mut i = 0;
    let mut dbuf = [0u8; 10];

    if row == 0 {
        i = 1;
    } else {
        while row > 0 {
            dbuf[i] = (row % 10) as u8;
            row /= 10;

            i += 1;
        }
    }

    // reverse order
    let mut j = i;
    while j > 0 {
        buf.push((b'0' + dbuf[j - 1]) as char);
        j -= 1;
    }
}

/// Returns the spreadsheet column name.
pub fn colname(col: ucell) -> String {
    let mut col_str = String::new();
    push_colname(&mut col_str, col);
    col_str
}

/// Returns the spreadsheet row name
pub fn rowname(row: ucell) -> String {
    let mut row_str = String::new();
    push_rowname(&mut row_str, row);
    row_str
}

/// Creates a cell-reference for use in formulas.
pub fn cellref(row: ucell, col: ucell) -> String {
    CellRef::simple(row, col).to_formula()
}

/// Creates a cell-reference for use in formulas.
pub fn cellref_table(table: &str, row: ucell, col: ucell) -> String {
    let mut c = CellRef::simple(row, col);
    c.table = Some(table);
    c.to_formula()
}

/// Creates a cell-reference for use in formulas.
pub fn rangeref(row: ucell,
                col: ucell,
                row_to: ucell,
                col_to: ucell) -> String {
    CellRange::simple(row, col, row_to, col_to).to_formula()
}

/// Creates a cell-reference for use in formulas.
pub fn rangeref_table(table: &str,
                      row: ucell,
                      col: ucell,
                      row_to: ucell,
                      col_to: ucell) -> String {
    let mut c = CellRange::simple(row, col, row_to, col_to);
    c.from.table = Some(table);
    c.to.table = Some(table);
    c.to_formula()
}
