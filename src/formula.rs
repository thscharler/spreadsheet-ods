use crate::ucell;

/// Returns the spreadsheet column name.
pub fn colref(mut col: ucell) -> String {
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

/// Creates a cell-reference for use in formulas.
pub fn cellref(row: ucell, col: ucell) -> String {
    let mut cell = String::from("[.");
    cell.push_str(&colref(col));
    cell.push_str(&(row + 1).to_string());
    cell.push_str("]");

    cell
}

/// Creates a cell-reference for use in formulas.
pub fn rangeref(row: ucell, col: ucell, row_to: ucell, col_to: ucell) -> String {
    let mut cell = String::from("[.");
    cell.push_str(&colref(col));
    cell.push_str(&(row + 1).to_string());
    cell.push_str(":");
    cell.push_str(&colref(col_to));
    cell.push_str(&(row_to + 1).to_string());
    cell.push_str("]");

    cell
}