//!
//! Defines types for cell references.
//!

use std::convert::TryFrom;
use std::fmt::{Display, Formatter};

use crate::{ucell, OdsError};

/// Reference to a cell.
///
/// ```
/// use spreadsheet_ods::CellRef;
/// use std::convert::TryFrom;
///
/// let c1 = CellRef::local(5, 2);
/// let c2 = CellRef::remote("spreadsheet-2", 7, 4);
/// let c3 = CellRef::try_from(".A5");
/// ```
///
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct CellRef {
    table: Option<String>,
    row_abs: bool, /* Absolute ($) reference */
    row: ucell,
    col_abs: bool, /* Absolute ($) reference */
    col: ucell,
}

impl TryFrom<&str> for CellRef {
    type Error = OdsError;

    fn try_from(s: &str) -> Result<Self, Self::Error> {
        let mut pos = 0usize;
        parse_cellref(s, &mut pos)
    }
}

impl Display for CellRef {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), core::fmt::Error> {
        let mut buf = String::new();
        push_cellref(&mut buf, self);
        write!(f, "{}", buf)
    }
}

impl CellRef {
    pub fn new() -> Self {
        Self {
            table: None,
            row: 0,
            col: 0,
            row_abs: false,
            col_abs: false,
        }
    }

    /// Creates a cellref within the same table.
    pub fn local(row: ucell, col: ucell) -> Self {
        Self {
            table: None,
            row,
            row_abs: false,
            col,
            col_abs: false,
        }
    }

    /// Creates a cellref that references another table.
    pub fn remote<S: Into<String>>(table: S, row: ucell, col: ucell) -> Self {
        Self {
            table: Some(table.into()),
            row,
            row_abs: false,
            col,
            col_abs: false,
        }
    }

    /// Table name for references into other tables.
    pub fn set_table<S: Into<String>>(&mut self, table: S) {
        self.table = Some(table.into());
    }

    /// Table name for references into other tables.
    pub fn table(&self) -> Option<&String> {
        self.table.as_ref()
    }

    pub fn set_row(&mut self, row: ucell) {
        self.row = row;
    }

    pub fn row(&self) -> ucell {
        self.row
    }

    /// "$" row reference
    pub fn set_row_abs(&mut self, abs: bool) {
        self.row_abs = abs;
    }

    /// "$" row reference
    pub fn row_abs(&self) -> bool {
        self.row_abs
    }

    pub fn set_col(&mut self, col: ucell) {
        self.col = col;
    }

    pub fn col(&self) -> ucell {
        self.col
    }

    /// "$" column reference
    pub fn set_col_abs(&mut self, abs: bool) {
        self.col_abs = abs;
    }

    /// "$" column reference
    pub fn col_abs(&self) -> bool {
        self.col_abs
    }

    /// Returns a cell reference for a formula.
    pub fn to_formula(&self) -> String {
        let mut buf = String::new();
        buf.push('[');
        push_cellref(&mut buf, self);
        buf.push(']');

        buf
    }

    /// Makes this CellReference into an absolute reference.
    pub fn absolute(mut self) -> Self {
        self.col_abs = true;
        self.row_abs = true;
        self
    }

    /// Makes this CellReference into an absolute reference.
    /// The column remains relative, the row is fixed.
    pub fn absolute_row(mut self) -> Self {
        self.row_abs = true;
        self
    }

    /// Makes this CellReference into an absolute reference.
    /// The row remains relative, the column is fixed.
    pub fn absolute_col(mut self) -> Self {
        self.col_abs = true;
        self
    }
}

/// A cell-range.
///
/// As usual for a spreadsheet this is meant as inclusive from and to.
///
/// ```
/// use spreadsheet_ods::CellRange;
/// let r1 = CellRange::local(0, 0, 9, 9);
/// let r2 = CellRange::origin_span(5, 5, (3, 3));
/// ```
///
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct CellRange {
    table: Option<String>,
    row_abs: bool, /* Absolute ($) reference */
    row: ucell,
    col_abs: bool, /* Absolute ($) reference */
    col: ucell,
    to_row_abs: bool, /* Absolute ($) reference */
    to_row: ucell,
    to_col_abs: bool, /* Absolute ($) reference */
    to_col: ucell,
}

impl CellRange {
    pub fn new() -> Self {
        Self {
            table: None,
            row_abs: false,
            row: 0,
            col_abs: false,
            col: 0,
            to_row_abs: false,
            to_row: 0,
            to_col_abs: false,
            to_col: 0,
        }
    }

    /// Creates the cell range from from + to data.
    pub fn local(row: ucell, col: ucell, to_row: ucell, to_col: ucell) -> Self {
        assert!(row <= to_row);
        assert!(col <= to_col);
        Self {
            table: None,
            row_abs: false,
            row,
            col_abs: false,
            col,
            to_row_abs: false,
            to_row,
            to_col_abs: false,
            to_col,
        }
    }

    /// Creates the cell range from from + to data.
    pub fn remote<S: Into<String>>(
        table: S,
        row: ucell,
        col: ucell,
        to_row: ucell,
        to_col: ucell,
    ) -> Self {
        assert!(row <= to_row);
        assert!(col <= to_col);
        Self {
            table: Some(table.into()),
            row_abs: false,
            row,
            col_abs: false,
            col,
            to_row_abs: false,
            to_row,
            to_col_abs: false,
            to_col,
        }
    }

    /// Creates the cell range from origin + spanning data.
    pub fn origin_span(row: ucell, col: ucell, span: (ucell, ucell)) -> Self {
        assert!(span.0 > 0);
        assert!(span.1 > 0);
        Self {
            table: None,
            row_abs: false,
            row,
            col_abs: false,
            col,
            to_row_abs: false,
            to_row: row + span.0 - 1,
            to_col_abs: false,
            to_col: col + span.1 - 1,
        }
    }
    /// Table name for references into other tables.
    pub fn set_table<S: Into<String>>(&mut self, table: S) {
        self.table = Some(table.into());
    }

    /// Table name for references into other tables.
    pub fn table(&self) -> Option<&String> {
        self.table.as_ref()
    }

    pub fn set_row(&mut self, row: ucell) {
        self.row = row;
    }

    pub fn row(&self) -> ucell {
        self.row
    }

    /// "$" row reference
    pub fn set_row_abs(&mut self, abs: bool) {
        self.row_abs = abs;
    }

    /// "$" row reference
    pub fn row_abs(&self) -> bool {
        self.row_abs
    }

    pub fn set_col(&mut self, col: ucell) {
        self.col = col;
    }

    pub fn col(&self) -> ucell {
        self.col
    }

    /// "$" column reference
    pub fn set_col_abs(&mut self, abs: bool) {
        self.col_abs = abs;
    }

    /// "$" column reference
    pub fn col_abs(&self) -> bool {
        self.col_abs
    }

    pub fn set_to_row(&mut self, to_row: ucell) {
        self.to_row = to_row;
    }

    pub fn to_row(&self) -> ucell {
        self.to_row
    }

    /// "$" row reference
    pub fn set_to_row_abs(&mut self, abs: bool) {
        self.to_row_abs = abs;
    }

    /// "$" row reference
    pub fn to_row_abs(&self) -> bool {
        self.to_row_abs
    }

    pub fn set_to_col(&mut self, to_col: ucell) {
        self.to_col = to_col;
    }

    pub fn to_col(&self) -> ucell {
        self.to_col
    }

    /// "$" column reference
    pub fn set_to_col_abs(&mut self, abs: bool) {
        self.to_col_abs = abs;
    }

    /// "$" column reference
    pub fn to_col_abs(&self) -> bool {
        self.to_col_abs
    }

    /// Returns a range reference for a formula.
    pub fn to_formula(&self) -> String {
        let mut buf = String::new();
        buf.push('[');
        push_cellrange(&mut buf, self);
        buf.push(']');
        buf
    }

    /// Makes this CellReference into an absolute reference.
    pub fn absolute(mut self) -> Self {
        self.col_abs = true;
        self.row_abs = true;
        self.to_col_abs = true;
        self.to_row_abs = true;
        self
    }

    /// Makes this CellReference into an absolute reference.
    /// The columns remain relative, the rows are fixed.
    pub fn absolute_rows(mut self) -> Self {
        self.row_abs = true;
        self.to_row_abs = true;
        self
    }

    /// Makes this CellReference into an absolute reference.
    /// The rows remain relative, the columns are fixed.
    pub fn absolute_cols(mut self) -> Self {
        self.col_abs = true;
        self.to_col_abs = true;
        self
    }

    /// Does the range contain the cell.
    /// This is inclusive for to_row and to_col!
    pub fn contains(&self, row: ucell, col: ucell) -> bool {
        row >= self.row && row <= self.to_row && col >= self.col && col <= self.to_col
    }

    /// Is this range any longer relevant, when looping rows first, then columns?
    pub fn out_looped(&self, row: ucell, col: ucell) -> bool {
        row > self.to_row || row == self.to_row && col > self.to_col
    }
}

impl TryFrom<&str> for CellRange {
    type Error = OdsError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        let mut pos = 0usize;
        parse_cellrange(value, &mut pos)
    }
}

impl Display for CellRange {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), core::fmt::Error> {
        let mut buf = String::new();
        push_cellrange(&mut buf, self);
        write!(f, "{}", buf)
    }
}

/// A range over columns.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct ColRange {
    col: ucell,
    to_col: ucell,
}

impl Display for ColRange {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        write!(f, "{}:{}", self.col, self.to_col)
    }
}

impl ColRange {
    pub fn new(col: ucell, to_col: ucell) -> Self {
        assert!(col <= to_col);
        Self { col, to_col }
    }

    pub fn set_col(&mut self, col: ucell) {
        self.col = col;
    }

    pub fn col(&self) -> ucell {
        self.col
    }

    pub fn set_to_col(&mut self, to_col: ucell) {
        self.to_col = to_col;
    }

    pub fn to_col(&self) -> ucell {
        self.to_col
    }

    /// Is the column in this range.
    /// The range is inclusive with the to_col.
    pub fn contains(&self, col: ucell) -> bool {
        col >= self.col && col <= self.to_col
    }
}

/// A range over rows.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct RowRange {
    pub row: ucell,
    pub to_row: ucell,
}

impl Display for RowRange {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        write!(f, "{}:{}", self.row, self.to_row)
    }
}

impl RowRange {
    pub fn new(row: ucell, to_row: ucell) -> Self {
        assert!(row <= to_row);
        Self { row, to_row }
    }

    /// Is the row in this range.
    /// The range is inclusive with the to_row.
    pub fn contains(&self, row: ucell) -> bool {
        row >= self.row && row <= self.to_row
    }
}

/// Parse the colname.
/// Stops when the colname ends and returns the byte position in end.
#[allow(clippy::manual_range_contains)]
pub(crate) fn parse_colname(buf: &str, pos: &mut usize) -> Option<ucell> {
    let mut col = 0u32;

    let mut loop_break = false;
    for (p, c) in buf[*pos..].char_indices() {
        if c < 'A' || c > 'Z' {
            loop_break = true;
            *pos += p;
            break;
        }

        let mut v = c as u32 - b'A' as u32;
        if v == 25 {
            v = 0;
            col = (col + 1) * 26;
        } else {
            v += 1;
            col *= 26;
        }
        col += v as u32;
    }
    // consumed all chars
    if !loop_break {
        *pos = buf.len();
    }

    if col == 0 {
        None
    } else {
        Some(col - 1)
    }
}

/// Parse the rowname.
/// Stops when the rowname ends and returns the byte position in end.
#[allow(clippy::manual_range_contains)]
pub(crate) fn parse_rowname(buf: &str, pos: &mut usize) -> Option<ucell> {
    let mut row = 0u32;

    let mut loop_break = false;
    for (p, c) in buf[*pos..].char_indices() {
        if c < '0' || c > '9' {
            loop_break = true;
            *pos += p;
            break;
        }

        row *= 10;
        row += c as u32 - '0' as u32;
    }

    // consumed all chars
    if !loop_break {
        *pos = buf.len();
    }

    if row == 0 {
        None
    } else {
        Some(row - 1)
    }
}

/// Parse a table-name in a reference
#[allow(clippy::collapsible_else_if)]
pub(crate) fn parse_tablename(buf: &str, pos: &mut usize) -> Result<Option<String>, OdsError> {
    let mut dot_idx = None;
    let mut any_quote = false;
    let mut state_quote = false;

    let abs_sheet = buf[*pos..].starts_with('$');
    if abs_sheet {
        *pos += 1;
    }

    for (p, c) in buf[*pos..].char_indices() {
        if !state_quote {
            if c == '\'' {
                state_quote = true;
                any_quote = true;
            }
            if c == '.' {
                dot_idx = Some(*pos + p);
                break;
            }
        } else {
            if c == '\'' {
                state_quote = false;
            }
        }
    }
    if dot_idx.is_none() {
        return Err(OdsError::Ods(format!(
            "No '.' in the cell reference {}",
            &buf[*pos..]
        )));
    }
    let dot_idx = dot_idx.unwrap();

    // Tablename
    let table = if dot_idx > *pos {
        if any_quote {
            // quoting rules: enclose with ' and double contained ''
            Some(buf[*pos..dot_idx].trim_matches('\'').replace("''", "'"))
        } else {
            Some(buf[*pos..dot_idx].to_string())
        }
    } else {
        None
    };

    *pos = dot_idx + 1;

    Ok(table)
}

/// Parse a cell reference.
pub(crate) fn parse_cellref(buf: &str, pos: &mut usize) -> Result<CellRef, OdsError> {
    let table = parse_tablename(buf, pos)?;

    let abs_col = buf[*pos..].starts_with('$');
    if abs_col {
        *pos += 1;
    }

    let col = parse_colname(buf, pos);
    if col.is_none() {
        return Err(OdsError::Ods(format!(
            "No colname in the cell reference {}",
            &buf[*pos..]
        )));
    }

    let abs_row = buf[*pos..].starts_with('$');
    if abs_row {
        *pos += 1;
    }

    let row = parse_rowname(buf, pos);
    if row.is_none() {
        return Err(OdsError::Ods(format!(
            "No rowname in the cell reference {}",
            &buf[*pos..]
        )));
    }

    Ok(CellRef {
        table,
        row: row.unwrap(),
        row_abs: abs_row,
        col: col.unwrap(),
        col_abs: abs_col,
    })
}

/// Parse a range ref.
pub(crate) fn parse_cellrange(buf: &str, pos: &mut usize) -> Result<CellRange, OdsError> {
    let cell = {
        let table = parse_tablename(buf, pos)?;

        let abs_col = buf[*pos..].starts_with('$');
        if abs_col {
            *pos += 1;
        }

        let col = match parse_colname(buf, pos) {
            None => {
                return Err(OdsError::Ods(format!(
                    "No colname in the cell reference {}",
                    &buf[*pos..]
                )))
            }
            Some(col) => col,
        };

        let abs_row = buf[*pos..].starts_with('$');
        if abs_row {
            *pos += 1;
        }

        let row = match parse_rowname(buf, pos) {
            None => {
                return Err(OdsError::Ods(format!(
                    "No rowname in the cell reference {}",
                    &buf[*pos..]
                )))
            }
            Some(row) => row,
        };

        (table, abs_row, row, abs_col, col)
    };

    // a range can be a single cell too
    let colon = buf[*pos..].starts_with(':');

    let to_cell = if colon {
        *pos += 1;

        let to_table = parse_tablename(buf, pos)?;

        let abs_to_col = buf[*pos..].starts_with('$');
        if abs_to_col {
            *pos += 1;
        }

        let to_col = match parse_colname(buf, pos) {
            None => {
                return Err(OdsError::Ods(format!(
                    "No colname in the cell reference {}",
                    &buf[*pos..]
                )))
            }
            Some(col) => col,
        };

        let abs_to_row = buf[*pos..].starts_with('$');
        if abs_to_row {
            *pos += 1;
        }

        let to_row = match parse_rowname(buf, pos) {
            None => {
                return Err(OdsError::Ods(format!(
                    "No rowname in the cell reference {}",
                    &buf[*pos..]
                )))
            }
            Some(row) => row,
        };

        (to_table, abs_to_row, to_row, abs_to_col, to_col)
    } else {
        cell.clone()
    };

    Ok(CellRange {
        table: cell.0,
        row_abs: cell.1,
        row: cell.2,
        col_abs: cell.3,
        col: cell.4,
        // to_table is ignored. should be the same as table.
        to_row_abs: to_cell.1,
        to_row: to_cell.2,
        to_col_abs: to_cell.3,
        to_col: to_cell.4,
    })
}

/// Parse a list of range refs
pub(crate) fn parse_cellranges(
    buf: &str,
    pos: &mut usize,
) -> Result<Option<Vec<CellRange>>, OdsError> {
    let mut v = None;

    loop {
        let r = parse_cellrange(buf, pos)?;

        if v.is_none() {
            v = Some(Vec::new());
        }
        if let Some(ref mut v) = v {
            v.push(r);
        }

        if *pos == buf.len() {
            break;
        }

        if !buf[*pos..].starts_with(' ') {
            return Err(OdsError::Ods(format!(
                "No blank between cellranges {}",
                &buf[*pos..]
            )));
        } else {
            *pos += 1;
        }
    }

    Ok(v)
}

/// Appends the spreadsheet column name.
pub(crate) fn push_colname(buf: &mut String, mut col: ucell) {
    let mut i = 0;
    let mut dbuf = [0u8; 7];

    if col == ucell::max_value() {
        // unroll first loop because of overflow
        dbuf[0] = 21;
        i += 1;
        col /= 26;
    } else {
        col += 1;
    }

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

/// Appends the spreadsheet row name
pub(crate) fn push_rowname(buf: &mut String, row: ucell) {
    let mut i = 0;
    let mut dbuf = [0u8; 10];

    // temp solution
    let mut row: u64 = row.into();
    row += 1;
    while row > 0 {
        dbuf[i] = (row % 10) as u8;
        row /= 10;

        i += 1;
    }

    // reverse order
    let mut j = i;
    while j > 0 {
        buf.push((b'0' + dbuf[j - 1]) as char);
        j -= 1;
    }
}

/// Appends the table-name
pub(crate) fn push_tablename(buf: &mut String, table: Option<&String>) {
    if let Some(table) = table {
        if table.contains(|c| c == '\'' || c == ' ' || c == '.') {
            buf.push('\'');
            buf.push_str(&table.replace('\'', "''"));
            buf.push('\'');
        } else {
            buf.push_str(table);
        }
        buf.push('.');
    } else {
        buf.push('.');
    }
}

/// Appends the cell reference
pub(crate) fn push_cellref(buf: &mut String, cellref: &CellRef) {
    push_tablename(buf, cellref.table.as_ref());
    if cellref.col_abs {
        buf.push('$');
    }
    push_colname(buf, cellref.col);
    if cellref.row_abs {
        buf.push('$');
    }
    push_rowname(buf, cellref.row);
}

/// Appends the range reference
pub(crate) fn push_cellrange(buf: &mut String, cellrange: &CellRange) {
    push_tablename(buf, cellrange.table.as_ref());
    if cellrange.col_abs {
        buf.push('$');
    }
    push_colname(buf, cellrange.col);
    if cellrange.row_abs {
        buf.push('$');
    }
    push_rowname(buf, cellrange.row);
    buf.push(':');
    buf.push('.');
    if cellrange.to_col_abs {
        buf.push('$');
    }
    push_colname(buf, cellrange.to_col);
    if cellrange.to_row_abs {
        buf.push('$');
    }
    push_rowname(buf, cellrange.to_row);
}

/// Returns a list of ranges as string.
pub(crate) fn cellranges_string(v: &[CellRange]) -> String {
    let mut buf = String::new();

    let mut first = true;
    for r in v {
        if first {
            first = false;
        } else {
            buf.push(' ');
        }
        push_cellrange(&mut buf, r);
    }

    buf
}

#[cfg(test)]
mod tests {
    use crate::refs::{
        parse_cellrange, parse_cellranges, parse_cellref, parse_colname, parse_rowname,
        push_cellrange, push_cellref, push_colname, push_rowname, push_tablename,
    };
    use crate::{ucell, CellRange, CellRef, OdsError};

    #[test]
    fn test_names() {
        let mut buf = String::new();

        push_colname(&mut buf, 0);
        assert_eq!(buf, "A");
        buf.clear();

        push_colname(&mut buf, 1);
        assert_eq!(buf, "B");
        buf.clear();

        push_colname(&mut buf, 26);
        assert_eq!(buf, "AA");
        buf.clear();

        push_colname(&mut buf, 675);
        assert_eq!(buf, "YZ");
        buf.clear();

        push_colname(&mut buf, 676);
        assert_eq!(buf, "ZA");
        buf.clear();

        push_colname(&mut buf, ucell::max_value() - 1);
        assert_eq!(buf, "MWLQKWU");
        buf.clear();

        push_colname(&mut buf, ucell::max_value());
        assert_eq!(buf, "MWLQKWV");
        buf.clear();

        push_rowname(&mut buf, 0);
        assert_eq!(buf, "1");
        buf.clear();

        push_rowname(&mut buf, 927);
        assert_eq!(buf, "928");
        buf.clear();

        push_rowname(&mut buf, ucell::max_value() - 1);
        assert_eq!(buf, "4294967295");
        buf.clear();

        push_rowname(&mut buf, ucell::max_value());
        assert_eq!(buf, "4294967296");
        buf.clear();

        push_tablename(&mut buf, Some(&"fable".to_string()));
        assert_eq!(buf, "fable.");
        buf.clear();

        push_tablename(&mut buf, Some(&"fa le".to_string()));
        assert_eq!(buf, "'fa le'.");
        buf.clear();

        push_tablename(&mut buf, Some(&"fa'le".to_string()));
        assert_eq!(buf, "'fa''le'.");
        buf.clear();

        push_tablename(&mut buf, Some(&"fa.le".to_string()));
        assert_eq!(buf, "'fa.le'.");
        buf.clear();

        push_tablename(&mut buf, None);
        assert_eq!(buf, ".");
        buf.clear();

        push_cellref(&mut buf, &CellRef::local(5, 6));
        assert_eq!(buf, ".G6");
        buf.clear();

        push_cellrange(&mut buf, &CellRange::local(5, 6, 7, 8));
        assert_eq!(buf, ".G6:.I8");
        buf.clear();

        push_cellrange(&mut buf, &CellRange::remote("blame", 5, 6, 7, 8));
        assert_eq!(buf, "blame.G6:.I8");
        buf.clear();
    }

    #[test]
    fn test_parse() -> Result<(), OdsError> {
        fn rowname(row: ucell) -> String {
            let mut row_str = String::new();
            push_rowname(&mut row_str, row);
            row_str
        }
        fn colname(col: ucell) -> String {
            let mut col_str = String::new();
            push_colname(&mut col_str, col);
            col_str
        }

        for i in 0..704 {
            let mut pos = 0usize;
            let cn = colname(i);
            let ccc = parse_colname(&cn, &mut pos);
            assert_eq!(Some(i), ccc);
            assert_eq!(cn.len(), pos);
        }

        for i in 0..101 {
            let mut pos = 0usize;
            let cn = rowname(i);
            let cr = parse_rowname(&cn, &mut pos);
            assert_eq!(Some(i), cr);
            assert_eq!(cn.len(), pos);
        }

        let mut pos = 0usize;
        let cn = "A32";
        let cc = parse_colname(cn, &mut pos);
        assert_eq!(Some(0), cc);
        assert_eq!(1, pos);

        let mut pos = 0usize;
        let cn = "AAAA32 ";
        let cc = parse_colname(cn, &mut pos);
        assert_eq!(Some(18278), cc);
        assert_eq!(4, pos);
        let cr = parse_rowname(cn, &mut pos);
        assert_eq!(Some(31), cr);
        assert_eq!(6, pos);

        let mut pos = 0usize;
        let cn = ".A3";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(cr, CellRef::local(2, 0));

        let mut pos = 0usize;
        let cn = ".$A3";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRef {
                table: None,
                row: 2,
                row_abs: false,
                col: 0,
                col_abs: true,
            }
        );

        let mut pos = 0usize;
        let cn = ".A$3";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRef {
                table: None,
                row: 2,
                row_abs: true,
                col: 0,
                col_abs: false,
            }
        );

        let mut pos = 0usize;
        let cn = "fufufu.A3";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRef {
                table: Some("fufufu".to_string()),
                row: 2,
                row_abs: false,
                col: 0,
                col_abs: false,
            }
        );

        let mut pos = 0usize;
        let cn = "'lak.moi'.A3";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRef {
                table: Some("lak.moi".to_string()),
                row: 2,
                row_abs: false,
                col: 0,
                col_abs: false,
            }
        );

        let mut pos = 0usize;
        let cn = "'lak''moi'.A3";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRef {
                table: Some("lak'moi".to_string()),
                row: 2,
                row_abs: false,
                col: 0,
                col_abs: false,
            }
        );

        let mut pos = 4usize;
        let cn = "****.B4";
        let cr = parse_cellref(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRef {
                table: None,
                row: 3,
                row_abs: false,
                col: 1,
                col_abs: false,
            }
        );

        let mut pos = 0usize;
        let cn = ".A3:.F9";
        let cr = parse_cellrange(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRange {
                table: None,
                row_abs: false,
                row: 2,
                col_abs: false,
                col: 0,
                to_row_abs: false,
                to_row: 8,
                to_col_abs: false,
                to_col: 5,
            }
        );

        let mut pos = 0usize;
        let cn = "table.A3:.F9";
        let cr = parse_cellrange(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRange {
                table: Some("table".to_string()),
                row_abs: false,
                row: 2,
                col_abs: false,
                col: 0,
                to_row_abs: false,
                to_row: 8,
                to_col_abs: false,
                to_col: 5,
            }
        );

        let mut pos = 0usize;
        let cn = "table.A3:.F9";
        let cr = parse_cellrange(cn, &mut pos)?;
        assert_eq!(
            cr,
            CellRange {
                table: Some("table".to_string()),
                row_abs: false,
                row: 2,
                col_abs: false,
                col: 0,
                to_row_abs: false,
                to_row: 8,
                to_col_abs: false,
                to_col: 5,
            }
        );

        let mut pos = 0usize;
        let cn = "table.A3:.F9 table.A4:.F10";
        let cr = parse_cellranges(cn, &mut pos)?;
        assert_eq!(
            cr,
            Some(vec![
                CellRange::remote("table", 2, 0, 8, 5),
                CellRange::remote("table", 3, 0, 9, 5),
            ])
        );

        Ok(())
    }
}
