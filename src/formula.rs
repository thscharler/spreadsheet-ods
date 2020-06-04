use crate::ucell;

/// Returns the spreadsheet column name.
pub fn colname(mut col: ucell) -> String {
    let mut col_str = String::new();

    if col == 0 {
        col_str.insert(0, 'A');
    }
    while col > 0 {
        let digit = (col % 26) as u8;
        let cc = (b'A' + digit) as char;
        col_str.insert(0, cc);
        col /= 26;
    }

    col_str
}

/// Returns the spreadsheet row name
pub fn rowname(row: ucell) -> String {
    (row + 1).to_string()
}

/// Creates a cell-reference for use in formulas.
pub fn cellref(row: ucell, col: ucell) -> String {
    cellref_abs("", row, false, col, false)
}

/// Creates a cell-reference for use in formulas.
pub fn cellref_ext(tabname: &str, row: ucell, col: ucell) -> String {
    cellref_abs(tabname, row, false, col, false)
}

/// Creates an absolute cell-reference for use in formulas.
pub fn cellref_abs(tabname: &str,
                   row: ucell,
                   row_abs: bool,
                   col: ucell,
                   col_abs: bool) -> String {
    let mut cell = String::new();
    cell.push('[');
    cell.push_str(tabname);
    cell.push('.');
    if col_abs {
        cell.push('$');
    }
    cell.push_str(&colname(col));
    if row_abs {
        cell.push('$');
    }
    cell.push_str(&rowname(row));
    cell.push(']');

    cell
}

/// Creates a cell-reference for use in formulas.
pub fn rangeref(row: ucell,
                col: ucell,
                row_to: ucell,
                col_to: ucell) -> String {
    rangeref_ext("", row, col, row_to, col_to)
}

/// Creates a cell-reference for use in formulas.
pub fn rangeref_ext(tabname: &str,
                    row: ucell,
                    col: ucell,
                    row_to: ucell,
                    col_to: ucell) -> String {
    let mut cell = String::new();
    cell.push('[');
    cell.push_str(tabname);
    cell.push('.');
    cell.push_str(&colname(col));
    cell.push_str(&rowname(row));
    cell.push(':');
    cell.push_str(&colname(col_to));
    cell.push_str(&rowname(row_to));
    cell.push(']');

    cell
}
