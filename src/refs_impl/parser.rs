use crate::refs_impl::ast::{
    OFAst, OFCellRange, OFCellRef, OFCol, OFColRange, OFIri, OFRow, OFRowRange, OFSheetName,
};
use crate::refs_impl::error::LocateError;
use crate::refs_impl::error::OFCode::*;
use crate::refs_impl::tokens::eat_space;
use crate::refs_impl::{conv, map_err, tokens, ParseResult, Span};

/// Parses a space separated list of cell-ranges.
pub(crate) fn parse_cell_range_list<'s>(
    rest: Span<'s>,
) -> ParseResult<'s, Option<Vec<OFCellRange<'s>>>> {
    let mut vec = Vec::new();

    let mut rest_loop = rest;
    loop {
        rest_loop = match parse_cell_range(rest_loop) {
            Ok((rest1, cell_range)) => {
                vec.push(cell_range);
                rest1
            }
            Err(e) if e.code == OFCCellRange => break,
            Err(e) => return map_err(e, OFCUnexpected),
        };

        rest_loop = eat_space(rest_loop);
        if rest_loop.is_empty() {
            break;
        }
    }

    if vec.is_empty() {
        Ok((rest_loop, None))
    } else {
        Ok((rest_loop, Some(vec)))
    }
}

pub(crate) fn parse_cell_range<'s>(rest: Span<'s>) -> ParseResult<'s, OFCellRange<'s>> {
    let (rest, iri) = parse_iri(eat_space(rest))?;
    let (rest, table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _dot) = parse_dot(eat_space(rest))?;
    let (rest, col) = parse_col_term(eat_space(rest))?;
    let (rest, row) = parse_row_term(rest)?;

    let (rest, _colon) = parse_colon_term(eat_space(rest))?;

    let (rest, to_table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _to_dot) = parse_dot(eat_space(rest))?;
    let (rest, to_col) = parse_col_term(eat_space(rest))?;
    let (rest, to_row) = parse_row_term(rest)?;

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
    let (rest, iri) = parse_iri(eat_space(rest))?;
    let (rest, table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _dot) = parse_dot(eat_space(rest))?;
    let (rest, col) = parse_col_term(eat_space(rest))?;
    let (rest, row) = parse_row_term(rest)?;

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
    let (rest, iri) = parse_iri(eat_space(rest))?;
    let (rest, table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _dot) = parse_dot(eat_space(rest))?;
    let (rest, col) = parse_col_term(eat_space(rest))?;

    let (rest, _colon) = parse_colon_term(eat_space(rest))?;

    let (rest, to_table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _to_dot) = parse_dot(eat_space(rest))?;
    let (rest, to_col) = parse_col_term(eat_space(rest))?;

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
    let (rest, iri) = parse_iri(eat_space(rest))?;
    let (rest, table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _dot) = parse_dot(eat_space(rest))?;
    let (rest, row) = parse_row_term(eat_space(rest))?;

    let (rest, _colon) = parse_colon_term(eat_space(rest))?;

    let (rest, to_table) = parse_sheet_name(eat_space(rest))?;
    let (rest, _to_dot) = parse_dot(eat_space(rest))?;
    let (rest, to_row) = parse_row_term(eat_space(rest))?;

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
        Err(e) => map_err(e, OFCUnexpected),
    }
}

fn parse_sheet_name<'s>(rest: Span<'s>) -> ParseResult<'s, Option<OFSheetName<'s>>> {
    // QuotedSheetName ::= '$'? SingleQuoted "."
    let (rest, sheet_name) = match tokens::sheet_name(eat_space(rest)) {
        Ok((rest1, (abs, sheet_name))) => {
            let term = OFAst::sheet_name(
                conv::try_bool_from_abs_flag(abs),
                conv::unquote_single(sheet_name),
            );

            (rest1, Some(term))
        }
        Err(e) if e.code == OFCSingleQuoteStart => (rest, None),
        Err(e) if e.code == OFCSheetName => (rest, None),
        Err(e) if e.code == OFCString => return map_err(e, OFCSheetName),
        Err(e) if e.code == OFCSingleQuoteEnd => return map_err(e, OFCSheetName),
        Err(e) => return map_err(e, OFCUnexpected),
    };

    Ok((rest, sheet_name))
}

#[allow(unused)]
fn parse_dot<'s>(rest: Span<'s>) -> ParseResult<'s, Span<'s>> {
    // required dot
    let (rest, dot) = match tokens::dot(eat_space(rest)) {
        Ok((rest1, dot)) => (rest1, dot),
        Err(e) if e.code == OFCDot => return map_err(e, OFCDot),
        Err(e) => return map_err(e, OFCUnexpected),
    };

    Ok((rest, dot))
}

#[allow(unused)]
fn parse_col_term<'s>(rest: Span<'s>) -> ParseResult<'s, OFCol<'s>> {
    let (rest, col) = match tokens::col(rest) {
        Ok((rest, col)) => (rest, col),
        Err(e) if e.code == OFCAlpha => return map_err(e, OFCCol),
        Err(e) => return map_err(e, OFCUnexpected),
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
        Err(e) => return map_err(e, OFCUnexpected),
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
        Err(e) => return map_err(e, OFCUnexpected),
    };

    Ok((rest, ()))
}
