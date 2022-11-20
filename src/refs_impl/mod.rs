use crate::refs_impl::error::{OFCode, ParseOFError};
use nom::Offset;
use nom_locate::LocatedSpan;
use std::fmt::{Debug, Display, Formatter};
use std::marker::PhantomData;
use std::{fmt, slice};

mod parser;

// clones from openformula
#[allow(unreachable_pub)]
#[allow(unused)]
pub mod conv;
#[allow(unused)]
#[allow(unreachable_pub)]
pub mod error;
#[allow(unused)]
#[allow(unreachable_pub)]
pub mod format;
#[allow(unused)]
#[allow(unreachable_pub)]
pub mod tokens;

/// Input type.
pub(crate) type Span<'a> = LocatedSpan<&'a str>;

/// Result type.
pub(crate) type ParseResult<'s, O> = Result<(Span<'s>, O), ParseOFError<'s>>;

#[derive(PartialEq)]
#[allow(unused)]
pub(crate) enum OFAst<'a> {
    /// CellRef
    NodeCellRef(OFCellRef<'a>),
    /// CellRange
    NodeCellRange(OFCellRange<'a>),
    /// ColRange
    NodeColRange(OFColRange<'a>),
    /// RowRange
    NodeRowRange(OFRowRange<'a>),
}

//
// Functions that return some OFxxx
//
impl<'a> OFAst<'a> {
    /// Creates a OFIri
    pub(crate) fn iri(iri: String) -> OFIri<'a> {
        OFIri {
            iri,
            phantom: Default::default(),
        }
    }

    /// Creates a OFSheetName
    pub(crate) fn sheet_name(abs: bool, name: String) -> OFSheetName<'a> {
        OFSheetName {
            abs,
            name,
            phantom: Default::default(),
        }
    }

    /// Creates a OFRow
    pub(crate) fn row(abs: bool, row: u32) -> OFRow<'a> {
        OFRow {
            abs,
            row,
            phantom: Default::default(),
        }
    }

    /// Creates a OFCol
    pub(crate) fn col(abs: bool, col: u32) -> OFCol<'a> {
        OFCol {
            abs,
            col,
            phantom: Default::default(),
        }
    }
}

//
// Functions that return some OFxxx
//
impl<'a> OFAst<'a> {
    /// CellRef variant
    pub(crate) fn cell_ref(
        iri: Option<OFIri<'a>>,
        table: Option<OFSheetName<'a>>,
        row: OFRow<'a>,
        col: OFCol<'a>,
    ) -> OFAst<'a> {
        OFAst::NodeCellRef(OFCellRef {
            iri,
            table,
            row,
            col,
        })
    }

    /// CellRange variant
    pub(crate) fn cell_range(
        iri: Option<OFIri<'a>>,
        table: Option<OFSheetName<'a>>,
        row: OFRow<'a>,
        col: OFCol<'a>,
        to_table: Option<OFSheetName<'a>>,
        to_row: OFRow<'a>,
        to_col: OFCol<'a>,
    ) -> OFAst<'a> {
        OFAst::NodeCellRange(OFCellRange {
            iri,
            table,
            row,
            col,
            to_table,
            to_row,
            to_col,
        })
    }

    /// ColRange variant
    pub(crate) fn col_range(
        iri: Option<OFIri<'a>>,
        table: Option<OFSheetName<'a>>,
        col: OFCol<'a>,
        to_table: Option<OFSheetName<'a>>,
        to_col: OFCol<'a>,
    ) -> OFAst<'a> {
        OFAst::NodeColRange(OFColRange {
            iri,
            table,
            col,
            to_table,
            to_col,
        })
    }

    /// RowRange variant
    pub(crate) fn row_range(
        iri: Option<OFIri<'a>>,
        table: Option<OFSheetName<'a>>,
        row: OFRow<'a>,
        to_table: Option<OFSheetName<'a>>,
        to_row: OFRow<'a>,
    ) -> OFAst<'a> {
        OFAst::NodeRowRange(OFRowRange {
            iri,
            table,
            row,
            to_table,
            to_row,
        })
    }
}

// OFIri *****************************************************************

/// Represents an external source reference.
pub(crate) struct OFIri<'a> {
    pub(crate) iri: String,
    ///
    pub(crate) phantom: PhantomData<&'a str>,
}

impl<'a> Debug for OFIri<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        write!(f, "'{}'#", self.iri)
    }
}

impl<'a> Display for OFIri<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        write!(f, "'{}'#", self.iri)
    }
}

impl<'a> PartialEq for OFIri<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.iri == other.iri
    }
}

// OFSheetName ***********************************************************

/// Sheet name.
pub(crate) struct OFSheetName<'a> {
    /// Absolute reference.
    pub(crate) abs: bool,
    /// Sheet name.
    pub(crate) name: String,
    ///
    pub(crate) phantom: PhantomData<&'a str>,
}

impl<'a> Debug for OFSheetName<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if self.abs {
            write!(f, "$")?;
        }
        write!(f, "'{}'.", self.name)?;
        Ok(())
    }
}

impl<'a> Display for OFSheetName<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if self.abs {
            write!(f, "$")?;
        }
        write!(f, "'{}'.", self.name)?;
        Ok(())
    }
}

impl<'a> PartialEq for OFSheetName<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.abs == other.abs && self.name == other.name
    }
}

// OFRow *****************************************************************

/// Row data for any reference.
pub(crate) struct OFRow<'a> {
    /// Absolute flag
    pub(crate) abs: bool,
    /// Row
    pub(crate) row: u32,
    ///
    pub(crate) phantom: PhantomData<&'a str>,
}

impl<'a> Debug for OFRow<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        format::fmt_abs(f, self.abs)?;
        format::fmt_row_name(f, self.row)?;
        Ok(())
    }
}

impl<'a> Display for OFRow<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        format::fmt_abs(f, self.abs)?;
        format::fmt_row_name(f, self.row)?;
        Ok(())
    }
}

impl<'a> PartialEq for OFRow<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.abs == other.abs && self.row == other.row
    }
}

// OFCol *****************************************************************

/// Column data for any reference.
pub(crate) struct OFCol<'a> {
    /// Absolute flag
    pub(crate) abs: bool,
    /// Col
    pub(crate) col: u32,
    ///
    pub(crate) phantom: PhantomData<&'a str>,
}

impl<'a> Debug for OFCol<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        format::fmt_abs(f, self.abs)?;
        format::fmt_col_name(f, self.col)?;
        Ok(())
    }
}

impl<'a> Display for OFCol<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        format::fmt_abs(f, self.abs)?;
        format::fmt_col_name(f, self.col)?;
        Ok(())
    }
}

impl<'a> PartialEq for OFCol<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.abs == other.abs && self.col == other.col
    }
}

// OFCellRef *************************************************************

/// CellRef
pub(crate) struct OFCellRef<'a> {
    /// External source
    pub(crate) iri: Option<OFIri<'a>>,
    /// Sheet for reference.
    pub(crate) table: Option<OFSheetName<'a>>,
    /// Row
    pub(crate) row: OFRow<'a>,
    /// Col
    pub(crate) col: OFCol<'a>,
}

impl<'a> Debug for OFCellRef<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}{}", self.col, self.row)
    }
}

impl<'a> Display for OFCellRef<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}{}", self.col, self.row)
    }
}

impl<'a> PartialEq for OFCellRef<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.iri == other.iri
            && self.table == other.table
            && self.col == other.col
            && self.row == other.row
    }
}

// OFCellRange **********************************************************

/// CellRange
pub(crate) struct OFCellRange<'a> {
    pub(crate) iri: Option<OFIri<'a>>,
    pub(crate) table: Option<OFSheetName<'a>>,
    pub(crate) row: OFRow<'a>,
    pub(crate) col: OFCol<'a>,
    pub(crate) to_table: Option<OFSheetName<'a>>,
    pub(crate) to_row: OFRow<'a>,
    pub(crate) to_col: OFCol<'a>,
}

impl<'a> Debug for OFCellRange<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}{}", self.col, self.row)?;
        write!(f, ":")?;
        if let Some(to_table) = &self.to_table {
            write!(f, "{}", to_table)?;
        }
        write!(f, "{}{}", self.to_col, self.to_row)?;
        Ok(())
    }
}

impl<'a> Display for OFCellRange<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}{}", self.col, self.row)?;
        write!(f, ":")?;
        if let Some(to_table) = &self.to_table {
            write!(f, "{}", to_table)?;
        }
        write!(f, "{}{}", self.to_col, self.to_row)?;
        Ok(())
    }
}

impl<'a> PartialEq for OFCellRange<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.iri == other.iri
            && self.table == other.table
            && self.col == other.col
            && self.row == other.row
            && self.to_table == other.to_table
            && self.to_col == other.to_col
            && self.to_row == other.to_row
    }
}

// OFRowRange ************************************************************

/// RowRange
pub(crate) struct OFRowRange<'a> {
    /// External source
    pub(crate) iri: Option<OFIri<'a>>,
    /// Sheet for reference.
    pub(crate) table: Option<OFSheetName<'a>>,
    /// Row
    pub(crate) row: OFRow<'a>,
    /// Sheet for reference.
    pub(crate) to_table: Option<OFSheetName<'a>>,
    /// Row
    pub(crate) to_row: OFRow<'a>,
}

impl<'a> Debug for OFRowRange<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}", self.row)?;
        write!(f, ":")?;
        if let Some(to_table) = &self.to_table {
            write!(f, "{}", to_table)?;
        }
        write!(f, "{}", self.to_row)?;
        Ok(())
    }
}

impl<'a> Display for OFRowRange<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}", self.row)?;
        write!(f, ":")?;
        if let Some(to_table) = &self.to_table {
            write!(f, "{}", to_table)?;
        }
        write!(f, "{}", self.to_row)?;
        Ok(())
    }
}

impl<'a> PartialEq for OFRowRange<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.iri == other.iri
            && self.table == other.table
            && self.row == other.row
            && self.to_table == other.to_table
            && self.to_row == other.to_row
    }
}

// ColRange **************************************************************

/// ColRange
pub(crate) struct OFColRange<'a> {
    /// External source
    pub(crate) iri: Option<OFIri<'a>>,
    /// Sheet for reference.
    pub(crate) table: Option<OFSheetName<'a>>,
    /// Col
    pub(crate) col: OFCol<'a>,
    /// Sheet for reference.
    pub(crate) to_table: Option<OFSheetName<'a>>,
    /// Col
    pub(crate) to_col: OFCol<'a>,
}

impl<'a> Debug for OFColRange<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}", self.col)?;
        write!(f, ":")?;
        if let Some(to_table) = &self.to_table {
            write!(f, "{}", to_table)?;
        }
        write!(f, "{}", self.to_col)?;
        Ok(())
    }
}

impl<'a> Display for OFColRange<'a> {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        if let Some(iri) = &self.iri {
            write!(f, "{}", iri)?;
        }
        if let Some(table) = &self.table {
            write!(f, "{}", table)?;
        }
        write!(f, "{}", self.col)?;
        write!(f, ":")?;
        if let Some(to_table) = &self.to_table {
            write!(f, "{}", to_table)?;
        }
        write!(f, "{}", self.to_col)?;
        Ok(())
    }
}

impl<'a> PartialEq for OFColRange<'a> {
    fn eq(&self, other: &Self) -> bool {
        self.iri == other.iri
            && self.table == other.table
            && self.col == other.col
            && self.to_table == other.to_table
            && self.to_col == other.to_col
    }
}

// Returns a new Span that reaches from the beginning of span0 to the end of span1.
//
// If any of the following conditions are violated, the result is Undefined Behavior:
// * Both the starting and other pointer must be either in bounds or one byte past the end of the same allocated object.
//      Should be guaranteed if both were obtained from on ast run.
// * Both pointers must be derived from a pointer to the same object.
//      Should be guaranteed if both were obtained from on ast run.
// * The distance between the pointers, in bytes, cannot overflow an isize.
// * The distance being in bounds cannot rely on “wrapping around” the address space.
pub(crate) unsafe fn span_union<'a>(span0: Span<'a>, span1: Span<'a>) -> Span<'a> {
    let ptr = span0.as_ptr();
    // offset to the start of span1 and add the length of span1.
    let size = span0.offset(&span1) + span1.len();

    unsafe {
        // The size should be within the original allocation, if both spans are from
        // the same ast run. We must ensure that the ast run doesn't generate
        // Spans out of nothing that end in the ast.
        let slice = slice::from_raw_parts(ptr, size);
        // This is all from a str originally and we never got down to bytes.
        let str = std::str::from_utf8_unchecked(slice);

        // As span0 was ok the offset used here is ok too.
        Span::new_from_raw_offset(span0.location_offset(), span0.location_line(), str, ())
    }
}

pub(crate) fn map_err<'s, O>(mut err: ParseOFError<'s>, code: OFCode) -> ParseResult<'s, O> {
    // Translates the code with some exceptions.
    if err.code != OFCode::OFCNomError
        && err.code != OFCode::OFCNomFailure
        && err.code != OFCode::OFCUnexpected
        && err.code != OFCode::OFCParseIncomplete
    {
        err.code = code;
    }

    Err(err)
}

#[track_caller]
pub(crate) fn panic_parse<'s>(e: ParseOFError<'s>) -> ! {
    unreachable!("{}", e)
}

/// Helps with keeping tracks in the parsers.
///
/// This can be squeezed between the call to another parser and the ?-operator.
///
/// Makes sure the tracer can keep track of the complete parse call tree.
pub(crate) trait TrackParseResult<'s, 't, O> {
    /// Translates the error code and adds the standard expect value.
    /// Then tracks the error and marks the current function as finished.
    fn trace(self, code: OFCode) -> Self;
}

impl<'s, 't, O> TrackParseResult<'s, 't, O> for ParseResult<'s, O> {
    fn trace(self, code: OFCode) -> Self {
        match self {
            Ok(_) => self,
            Err(e) => map_err(e, code),
        }
    }
}
