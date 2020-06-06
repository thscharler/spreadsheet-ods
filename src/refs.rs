use crate::ucell;

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


/// Reference to a cell.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct CellRef {
    // Tablename
    pub table: Option<String>,
    // Row
    pub row: ucell,
    // Absolute ($) reference
    pub abs_row: bool,
    // Column
    pub col: ucell,
    // Absolute ($) reference
    pub abs_col: bool,
}

impl CellRef {
    pub fn simple(row: ucell, col: ucell) -> Self {
        Self {
            table: None,
            row,
            abs_row: false,
            col,
            abs_col: false,
        }
    }

    pub fn table<S: Into<String>>(table: S, row: ucell, col: ucell) -> Self {
        Self {
            table: Some(table.into()),
            row,
            abs_row: false,
            col,
            abs_col: false,
        }
    }

    /// Returns the spreadsheet column name.
    pub fn colname(&self) -> String {
        colname(self.col)
    }

    /// Returns the spreadsheet row name.
    pub fn rowname(&self) -> String {
        rowname(self.row)
    }

    /// Returns a cell reference.
    pub fn to_ref(&self) -> String {
        let mut refstr = String::new();
        if let Some(table) = &self.table {
            refstr.push_str(table);
        }
        refstr.push('.');
        if self.abs_col {
            refstr.push('$');
        }
        push_colname(&mut refstr, self.col);
        if self.abs_row {
            refstr.push('$');
        }
        push_rowname(&mut refstr, self.row);

        refstr
    }

    /// Returns a cell reference for a formula.
    pub fn to_formula(&self) -> String {
        let mut refstr = String::new();
        refstr.push('[');
        if let Some(table) = &self.table {
            refstr.push_str(table);
        }
        refstr.push('.');
        if self.abs_col {
            refstr.push('$');
        }
        push_colname(&mut refstr, self.col);
        if self.abs_row {
            refstr.push('$');
        }
        push_rowname(&mut refstr, self.row);
        refstr.push(']');

        refstr
    }
}

/// A cell-range.
/// As usual for a spreadsheet this is meant as inclusive from and to.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct CellRange {
    pub from: CellRef,
    pub to: CellRef,
}

impl CellRange {
    /// Creates the cell range from from + to data.
    pub fn simple(row: ucell, col: ucell, row_to: ucell, col_to: ucell) -> Self {
        assert!(row <= row_to);
        assert!(col <= col_to);
        Self {
            from: CellRef::simple(row, col),
            to: CellRef::simple(row_to, col_to),
        }
    }

    /// Creates the cell range from from + to data.
    pub fn table<S: Into<String>>(table: S, row: ucell, col: ucell, row_to: ucell, col_to: ucell) -> Self {
        assert!(row <= row_to);
        assert!(col <= col_to);
        let table = table.into();
        Self {
            from: CellRef::table(table.to_string(), row, col),
            to: CellRef::table(table.to_string(), row_to, col_to),
        }
    }

    /// Creates the cell range from origin + spanning data.
    pub fn origin_span(row: ucell, col: ucell, span: (ucell, ucell)) -> Self {
        assert!(span.0 > 0);
        assert!(span.1 > 0);
        Self {
            from: CellRef::simple(row, col),
            to: CellRef::simple(row + span.0 - 1, col + span.1 - 1),
        }
    }

    /// Returns a range reference.
    pub fn to_ref(&self) -> String {
        let mut refstr = String::new();
        refstr.push_str(&self.from.to_ref());
        refstr.push(':');
        refstr.push_str(&self.to.to_ref());
        refstr
    }

    /// Returns a range reference for a formula.
    pub fn to_formula(&self) -> String {
        let mut refstr = String::new();
        refstr.push('[');
        refstr.push_str(&self.from.to_ref());
        refstr.push(':');
        refstr.push_str(&self.to.to_ref());
        refstr.push(']');
        refstr
    }

    /// Does the range contain the cell.
    pub fn contains(&self, row: ucell, col: ucell) -> bool {
        row >= self.from.row && row <= self.to.row
            && col >= self.from.col && col <= self.to.col
    }

    /// Is this range any longer relevant, when looping rows first, then columns?
    pub fn out_looped(&self, row: ucell, col: ucell) -> bool {
        row > self.to.row
            || row == self.to.row && col > self.to.col
    }
}

/// A range over columns.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct ColRange {
    pub from: ucell,
    pub to: ucell,
}

impl ColRange {
    pub fn new(from: ucell, to: ucell) -> Self {
        assert!(from <= to);
        Self {
            from,
            to,
        }
    }

    pub fn contains(&self, col: ucell) -> bool {
        col >= self.from && col <= self.to
    }
}

/// A range over rows.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct RowRange {
    pub from: ucell,
    pub to: ucell,
}

impl RowRange {
    pub fn new(from: ucell, to: ucell) -> Self {
        assert!(from <= to);
        Self {
            from,
            to,
        }
    }

    pub fn contains(&self, row: ucell) -> bool {
        row >= self.from && row <= self.to
    }
}














