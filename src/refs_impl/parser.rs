use crate::refs_impl::ast::{
    OFAst, OFCellRange, OFCellRef, OFCol, OFColRange, OFIri, OFRow, OFRowRange, OFSheetName,
};
use crate::refs_impl::error::OFCode::*;
use crate::refs_impl::error::{LocateError, ParseOFError};
use crate::refs_impl::tokens::eat_space;
use crate::refs_impl::tokens::nomtokens::space;
use crate::refs_impl::{conv, map_err, panic_parse, tokens, ParseResult, Span, TrackParseResult};

/// Parses a space separated list of cell-ranges.
pub(crate) fn parse_cell_range_list<'s>(
    i: Span<'s>,
) -> ParseResult<'s, Option<Vec<OFCellRange<'s>>>> {
    let mut vec = Vec::new();

    let mut rest_loop = i;
    let rest2 = loop {
        let (rest1, cell_range) = parse_cell_range(rest_loop)?;
        vec.push(cell_range);

        if rest1.is_empty() {
            break rest1;
        }

        // eat one space or fail
        let rest1 = match space(rest1) {
            Ok((rest1, _sp)) => rest1,
            Err(e) => return Err(ParseOFError::nom(e)),
        };

        rest_loop = rest1;
    };

    if vec.is_empty() {
        Ok((rest2, None))
    } else {
        Ok((rest2, Some(vec)))
    }
}

pub(crate) fn parse_cell_range<'s>(rest: Span<'s>) -> ParseResult<'s, OFCellRange<'s>> {
    let (rest, iri) = parse_iri(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, table) = parse_sheet_name(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, _dot) = parse_dot(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, col) = parse_col_term(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, row) = parse_row_term(rest).trace(OFCCellRange)?;

    let (rest, _colon) = parse_colon_term(eat_space(rest)).trace(OFCCellRange)?;

    let (rest, to_table) = parse_sheet_name(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, _to_dot) = parse_dot(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, to_col) = parse_col_term(eat_space(rest)).trace(OFCCellRange)?;
    let (rest, to_row) = parse_row_term(rest).trace(OFCCellRange)?;

    let ast = OFCellRange {
        iri,
        table,
        row,
        col,
        to_table,
        to_row,
        to_col,
    };

    Ok((rest, ast))
}

pub(crate) fn parse_cell_ref<'s>(rest: Span<'s>) -> ParseResult<'s, OFCellRef<'s>> {
    let (rest, iri) = parse_iri(eat_space(rest)).trace(OFCCellRef)?;
    let (rest, table) = parse_sheet_name(eat_space(rest)).trace(OFCCellRef)?;
    let (rest, _dot) = parse_dot(eat_space(rest)).trace(OFCCellRef)?;
    let (rest, col) = parse_col_term(eat_space(rest)).trace(OFCCellRef)?;
    let (rest, row) = parse_row_term(rest).trace(OFCCellRef)?;

    let ast = OFCellRef {
        iri,
        table,
        row,
        col,
    };

    Ok((rest, ast))
}

#[allow(unused)]
pub(crate) fn parse_col_range<'s>(rest: Span<'s>) -> ParseResult<'s, OFColRange<'s>> {
    let (rest, iri) = parse_iri(eat_space(rest)).trace(OFCColRange)?;
    let (rest, table) = parse_sheet_name(eat_space(rest)).trace(OFCColRange)?;
    let (rest, _dot) = parse_dot(eat_space(rest)).trace(OFCColRange)?;
    let (rest, col) = parse_col_term(eat_space(rest)).trace(OFCColRange)?;

    let (rest, _colon) = parse_colon_term(eat_space(rest)).trace(OFCColRange)?;

    let (rest, to_table) = parse_sheet_name(eat_space(rest)).trace(OFCColRange)?;
    let (rest, _to_dot) = parse_dot(eat_space(rest)).trace(OFCColRange)?;
    let (rest, to_col) = parse_col_term(eat_space(rest)).trace(OFCColRange)?;

    let ast = OFColRange {
        iri,
        table,
        col,
        to_table,
        to_col,
    };

    Ok((rest, ast))
}

#[allow(unused)]
pub(crate) fn parse_row_range<'s>(rest: Span<'s>) -> ParseResult<'s, OFRowRange<'s>> {
    let (rest, iri) = parse_iri(eat_space(rest)).trace(OFCRowRange)?;
    let (rest, table) = parse_sheet_name(eat_space(rest)).trace(OFCRowRange)?;
    let (rest, _dot) = parse_dot(eat_space(rest)).trace(OFCRowRange)?;
    let (rest, row) = parse_row_term(eat_space(rest)).trace(OFCRowRange)?;

    let (rest, _colon) = parse_colon_term(eat_space(rest)).trace(OFCRowRange)?;

    let (rest, to_table) = parse_sheet_name(eat_space(rest)).trace(OFCRowRange)?;
    let (rest, _to_dot) = parse_dot(eat_space(rest)).trace(OFCRowRange)?;
    let (rest, to_row) = parse_row_term(eat_space(rest)).trace(OFCRowRange)?;

    let ast = OFRowRange {
        iri,
        table,
        row,
        to_table,
        to_row,
    };

    Ok((rest, ast))
}

#[allow(unused)]
fn parse_iri<'s>(rest: Span<'s>) -> ParseResult<'s, Option<OFIri<'s>>> {
    // (IRI '#')?
    match tokens::iri(eat_space(rest)) {
        Ok((rest1, iri)) => {
            let term = OFAst::iri(conv::unquote_single(iri));
            Ok((rest1, Some(term)))
        }
        // Fail to start any of these
        Err(e) if e.code == OFCSingleQuoteStart || e.code == OFCHashtag => Ok((rest, None)),
        Err(e) if e.code == OFCString => map_err(e, OFCIri),
        Err(e) if e.code == OFCSingleQuoteEnd => map_err(e, OFCIri),
        Err(e) => panic_parse(e),
    }
}

fn parse_sheet_name<'s>(rest: Span<'s>) -> ParseResult<'s, Option<OFSheetName<'s>>> {
    // QuotedSheetName ::= '$'? SingleQuoted "."
    let (rest, sheet_name) = match tokens::quoted_sheet_name(eat_space(rest)) {
        Ok((rest1, (abs, sheet_name))) => {
            let term = OFAst::sheet_name(
                conv::try_bool_from_abs_flag(abs),
                conv::unquote_single(sheet_name),
            );

            (rest1, Some(term))
        }
        Err(e) if e.code == OFCSingleQuoteStart => (rest, None),
        Err(e) if e.code == OFCString => return map_err(e, OFCSheetName),
        Err(e) if e.code == OFCSingleQuoteEnd => return map_err(e, OFCSheetName),
        Err(e) => panic_parse(e),
    };

    Ok((rest, sheet_name))
}

#[allow(unused)]
fn parse_dot<'s>(rest: Span<'s>) -> ParseResult<'s, Span<'s>> {
    // required dot
    let (rest, dot) = match tokens::dot(eat_space(rest)) {
        Ok((rest1, dot)) => (rest1, dot),
        Err(e) if e.code == OFCDot => return map_err(e, OFCDot),
        Err(e) => panic_parse(e),
    };

    Ok((rest, dot))
}

#[allow(unused)]
fn parse_col_term<'s>(rest: Span<'s>) -> ParseResult<'s, OFCol<'s>> {
    let (rest, col) = match tokens::col(rest) {
        Ok((rest, col)) => (rest, col),
        Err(e) if e.code == OFCAlpha => return map_err(e, OFCCol),
        Err(e) => panic_parse(e),
    };

    let col = OFAst::col(
        conv::try_bool_from_abs_flag(col.0),
        conv::try_u32_from_colname(col.1).locate_err(rest)?,
    );

    Ok((rest, col))
}

#[allow(unused)]
fn parse_row_term<'s>(rest: Span<'s>) -> ParseResult<'s, OFRow<'s>> {
    let (rest, row) = match tokens::row(rest) {
        Ok((rest, row)) => (rest, row),
        Err(e) if e.code == OFCDigit => return map_err(e, OFCRow),
        Err(e) => panic_parse(e),
    };

    let row = OFAst::row(
        conv::try_bool_from_abs_flag(row.0),
        conv::try_u32_from_rowname(row.1).locate_err(rest)?,
    );

    Ok((rest, row))
}

fn parse_colon_term<'s>(rest: Span<'s>) -> ParseResult<'s, ()> {
    // required dot
    let (rest, _colon) = match tokens::colon(eat_space(rest)) {
        Ok((rest1, dot)) => (rest1, dot),
        Err(e) if e.code == OFCColon => return map_err(e, OFCColon),
        Err(e) => panic_parse(e),
    };

    Ok((rest, ()))
}
