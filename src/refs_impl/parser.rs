use crate::refs_impl::parser::parser::{parse_col, parse_iri, parse_row, parse_sheet_name};
use crate::refs_impl::parser::tokens::{colon, dot};
use crate::{CellRange, CellRef, ColRange, RowRange};
use kparse::prelude::*;
use kparse::tracker::TrackResult;
#[cfg(debug_assertions)]
use kparse::tracker::TrackSpan;
use kparse::{Code, Context, ParserError};
use nom::character::complete::{multispace0, multispace1};
use std::fmt::{Display, Formatter};
use CRCode::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CRCode {
    CRNomError,

    CRCellRangeList,
    CRCellRange,
    CRColRange,
    CRRowRange,
    CRCellRef,

    CRIri,
    CRSheetName,

    CRCol,
    CRColInteger,
    CRColon,
    CRDollar,
    CRDot,
    CRHash,
    CRRow,
    CRRowInteger,
    CRSingleQuoteEnd,
    CRSingleQuoteStart,
    CRString,
    CRUnquotedName,
}

impl Display for CRCode {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let str = match self {
            CRNomError => "Nom",
            CRCellRangeList => "cell-range list",
            CRCellRange => "cell-range",
            CRColRange => "col-range",
            CRRowRange => "row-range",
            CRCellRef => "cell-ref",
            CRIri => "iri",
            CRSheetName => "sheet-name",
            CRCol => "col",
            CRColInteger => "col int",
            CRColon => ":",
            CRDollar => "$",
            CRDot => ".",
            CRHash => "#",
            CRRow => "row",
            CRRowInteger => "row int",
            CRSingleQuoteEnd => "' start",
            CRSingleQuoteStart => "' end",
            CRString => "str",
            CRUnquotedName => "unquoted",
        };
        write!(f, "{}", str)
    }
}

impl Code for CRCode {
    const NOM_ERROR: Self = Self::CRNomError;
}

#[cfg(debug_assertions)]
pub(crate) type CSpan<'s> = TrackSpan<'s, CRCode, &'s str>;
#[cfg(not(debug_assertions))]
pub(crate) type CSpan<'s> = &'s str;
pub(crate) type CParserResult<'s, O> = TrackResult<CRCode, CSpan<'s>, O, ()>;
pub(crate) type CNomResult<'s> = TrackResult<CRCode, CSpan<'s>, CSpan<'s>, ()>;
pub(crate) type CParserError<'s> = ParserError<CRCode, CSpan<'s>, ()>;

pub(crate) fn parse_cell_ref(input: CSpan<'_>) -> CParserResult<'_, CellRef> {
    Context.enter(CRCellRef, input);

    let (rest, iri) = parse_iri(input).track()?;
    let (rest, table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_col, col)) = parse_col(rest).track()?;
    let (rest, (abs_row, row)) = parse_row(rest).track()?;

    Context.ok(
        rest,
        input,
        CellRef::new_all(iri, table, abs_row, row, abs_col, col),
    )
}

pub(crate) fn parse_cell_range_list(input: CSpan<'_>) -> CParserResult<'_, Option<Vec<CellRange>>> {
    Context.enter(CRCellRangeList, input);

    let mut vec = Vec::new();

    let mut rest_loop = input;
    loop {
        let rest = match parse_cell_range(rest_loop) {
            Ok((rest1, cell_range)) => {
                vec.push(cell_range);
                rest1
            }
            Err(nom::Err::Error(e)) if e.code == CRDot => {
                break;
            }
            Err(e) => return Context.err(e),
        };

        let (rest, _) = multispace0(rest)?;

        if rest.len() == 0 {
            break;
        }

        rest_loop = rest;
    }

    if vec.is_empty() {
        Context.ok(rest_loop, input, None)
    } else {
        Context.ok(rest_loop, input, Some(vec))
    }
}

pub(crate) fn parse_cell_range(input: CSpan<'_>) -> CParserResult<'_, CellRange> {
    Context.enter(CRCellRange, input);

    let (rest, iri) = parse_iri(input).track()?;
    let (rest, table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_col, col)) = parse_col(rest).track()?;
    let (rest, (abs_row, row)) = parse_row(rest).track()?;

    let (rest, _) = colon(rest).track()?;

    let (rest, to_table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_to_col, to_col)) = parse_col(rest).track()?;
    let (rest, (abs_to_row, to_row)) = parse_row(rest).track()?;

    Context.ok(
        rest,
        input,
        CellRange::new_all(
            iri, table, abs_row, row, abs_col, col, to_table, abs_to_row, to_row, abs_to_col,
            to_col,
        ),
    )
}

pub(crate) fn parse_col_range(input: CSpan<'_>) -> CParserResult<'_, ColRange> {
    Context.enter(CRColRange, input);

    let (rest, iri) = parse_iri(input).track()?;
    let (rest, table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_col, col)) = parse_col(rest).track()?;

    let (rest, _) = colon(rest).track()?;

    let (rest, to_table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_to_col, to_col)) = parse_col(rest).track()?;

    Context.ok(
        rest,
        input,
        ColRange::new_all(iri, table, abs_col, col, to_table, abs_to_col, to_col),
    )
}

pub(crate) fn parse_row_range(input: CSpan<'_>) -> CParserResult<'_, RowRange> {
    Context.enter(CRRowRange, input);

    let (rest, iri) = parse_iri(input).track()?;
    let (rest, table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_row, row)) = parse_row(rest).track()?;

    let (rest, _) = colon(rest).track()?;

    let (rest, to_table) = parse_sheet_name(rest).track()?;
    let (rest, _) = dot(rest).track()?;
    let (rest, (abs_to_row, to_row)) = parse_row(rest).track()?;

    Context.ok(
        rest,
        input,
        RowRange::new_all(iri, table, abs_row, row, to_table, abs_to_row, to_row),
    )
}

mod conv {
    use crate::refs_impl::parser::CSpan;
    #[cfg(not(debug_assertions))]
    use kparse::prelude::*;
    use std::error::Error;
    use std::fmt::{Display, Formatter};
    use std::num::IntErrorKind;
    use std::str::FromStr;

    /// Replaces two single quotes (') with a single on.
    /// Strips one leading and one trailing quote.
    pub(crate) fn unquote_single(i: CSpan<'_>) -> Result<String, ()> {
        let i = match i.strip_prefix('\'') {
            None => i.fragment(),
            Some(s) => s,
        };
        let i = match i.strip_suffix('\'') {
            None => i,
            Some(s) => s,
        };

        Ok(i.replace("''", "'"))
    }

    /// Parse a bool if a '$' exists.
    pub(crate) fn try_bool_from_abs_flag(i: Option<CSpan<'_>>) -> bool {
        if let Some(i) = i {
            *i.fragment() == "$"
        } else {
            false
        }
    }

    /// Error for try_u32_from_rowname.
    #[allow(variant_size_differences)]
    #[derive(Debug, Clone, PartialEq, Eq)]
    pub(crate) enum ParseRownameError {
        /// Value being parsed is empty.
        ///
        /// This variant will be constructed when parsing an empty string.
        Empty,
        /// Contains an invalid digit in its Context.
        ///
        /// Among other causes, this variant will be constructed when parsing a string that
        /// contains a non-ASCII char.
        ///
        /// This variant is also constructed when a `+` or `-` is misplaced within a string
        /// either on its own or in the middle of a number.
        InvalidDigit,
        /// Integer is too large to store in target integer type.
        PosOverflow,
        /// Integer is too small to store in target integer type.
        NegOverflow,
        /// Value was Zero
        ///
        /// This variant will be emitted when the parsing string has a value of zero, which
        /// would be illegal for non-zero types.
        Zero,
        /// Something else.
        Other,
    }

    impl Display for ParseRownameError {
        fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
            match self {
                ParseRownameError::Empty => write!(f, "Input was empty")?,
                ParseRownameError::InvalidDigit => write!(f, "Invalid digit")?,
                ParseRownameError::PosOverflow => write!(f, "Positive overflow")?,
                ParseRownameError::NegOverflow => write!(f, "Negative overflow")?,
                ParseRownameError::Zero => write!(f, "Zero")?,
                ParseRownameError::Other => write!(f, "Other")?,
            }
            Ok(())
        }
    }

    impl Error for ParseRownameError {}

    /// Parse a row number to a row index.
    #[allow(clippy::explicit_auto_deref)]
    pub(crate) fn try_u32_from_rowname(i: CSpan<'_>) -> Result<u32, ParseRownameError> {
        match u32::from_str(i.fragment()) {
            Ok(v) if v > 0 => Ok(v - 1),
            Ok(_v) => Err(ParseRownameError::Zero),
            Err(e) => Err(match e.kind() {
                IntErrorKind::Empty => ParseRownameError::Empty,
                IntErrorKind::InvalidDigit => ParseRownameError::InvalidDigit,
                IntErrorKind::PosOverflow => ParseRownameError::PosOverflow,
                IntErrorKind::NegOverflow => ParseRownameError::NegOverflow,
                IntErrorKind::Zero => ParseRownameError::Zero,
                _ => ParseRownameError::Other,
            }),
        }
    }

    /// Error for try_u32_from_colname.
    #[allow(variant_size_differences)]
    #[derive(Debug, Clone, PartialEq, Eq)]
    pub(crate) enum ParseColnameError {
        /// Invalid column character.
        InvalidChar,
        /// Invalid column name.
        InvalidColname,
    }

    impl Display for ParseColnameError {
        fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
            match self {
                ParseColnameError::InvalidChar => {
                    write!(f, "Invalid char")?;
                }
                ParseColnameError::InvalidColname => {
                    write!(f, "Invalid colname")?;
                }
            }
            Ok(())
        }
    }

    impl Error for ParseColnameError {}

    /// Parse a col label to a column index.
    pub(crate) fn try_u32_from_colname(i: CSpan<'_>) -> Result<u32, ParseColnameError> {
        let mut col = 0u32;

        for c in (*i).chars() {
            if !('A'..='Z').contains(&c) {
                return Err(ParseColnameError::InvalidChar);
            }

            let mut v = c as u32 - b'A' as u32;
            if v == 25 {
                v = 0;
                col = (col + 1) * 26;
            } else {
                v += 1;
                col *= 26;
            }
            col += v;
        }

        if col == 0 {
            Err(ParseColnameError::InvalidColname)
        } else {
            Ok(col - 1)
        }
    }
}

mod parser {
    use crate::refs_impl::parser::conv::{unquote_single, ParseColnameError, ParseRownameError};
    use crate::refs_impl::parser::tokens::{col, hashtag, row, sheet_name, single_quoted_string};
    use crate::refs_impl::parser::CRCode::*;
    use crate::refs_impl::parser::{conv, CParserError, CParserResult, CRCode, CSpan};
    use kparse::combinators::{track, transform, transform_p};
    use kparse::prelude::*;
    use nom::sequence::terminated;

    pub(crate) fn parse_iri(input: CSpan<'_>) -> CParserResult<'_, Option<String>> {
        let parsed = track(
            CRIri,
            terminated(
                transform(single_quoted_string, |v| unquote_single(v), CRIri),
                hashtag,
            ),
        )(input);

        let (rest, iri) = match parsed {
            Ok((rest, iri)) => (rest, Some(iri)),
            Err(nom::Err::Error(e)) if e.code == CRSingleQuoteStart => (input, None),
            Err(nom::Err::Error(e)) if e.code == CRHash => (input, None),
            Err(e) => return Err(e),
        };

        Ok((rest, iri))
    }

    pub(crate) fn parse_sheet_name(input: CSpan<'_>) -> CParserResult<'_, Option<String>> {
        let parsed = track(
            CRSheetName,
            transform(sheet_name, |v| unquote_single(v), CRSheetName),
        )(input);

        let (rest, name) = match parsed {
            Ok((rest, name)) => (rest, Some(name)),
            Err(nom::Err::Error(e)) if e.code == CRSingleQuoteStart || e.code == CRUnquotedName => {
                (input, None)
            }
            Err(e) => return Err(e),
        };

        Ok((rest, name))
    }

    impl<'s> WithSpan<CRCode, CSpan<'s>, CParserError<'s>> for ParseRownameError {
        fn with_span(self, code: CRCode, span: CSpan<'s>) -> nom::Err<CParserError<'s>> {
            nom::Err::Error(CParserError::new(code, span).with_cause(self))
        }
    }

    impl<'s> WithSpan<CRCode, CSpan<'s>, CParserError<'s>> for ParseColnameError {
        fn with_span(self, code: CRCode, span: CSpan<'s>) -> nom::Err<CParserError<'s>> {
            nom::Err::Error(CParserError::new(code, span).with_cause(self))
        }
    }

    pub(crate) fn parse_row(input: CSpan<'_>) -> CParserResult<'_, (bool, u32)> {
        track(
            CRRow,
            transform_p(row, |(abs, row)| {
                Ok((
                    conv::try_bool_from_abs_flag(abs),
                    conv::try_u32_from_rowname(row).with_span(CRRowInteger, row)?,
                ))
            }),
        )(input)
    }

    pub(crate) fn parse_col(input: CSpan<'_>) -> CParserResult<'_, (bool, u32)> {
        track(
            CRCol,
            transform_p(col, |(abs, col)| {
                Ok((
                    conv::try_bool_from_abs_flag(abs),
                    conv::try_u32_from_colname(col).with_span(CRColInteger, col)?,
                ))
            }),
        )(input)
    }
}

mod tokens {
    use crate::refs_impl::parser::CRCode::*;
    use crate::refs_impl::parser::{CNomResult, CParserError, CParserResult, CSpan};
    use kparse::combinators::error_code;
    use kparse::prelude::*;
    use nom::branch::alt;
    use nom::bytes::complete::{tag, take_while1};
    use nom::character::complete::{alpha1, char as nchar, digit1, none_of};
    use nom::combinator::{opt, recognize};
    use nom::multi::{count, many0, many1};
    use nom::sequence::{preceded, tuple};

    const SINGLE_QUOTE: char = '\'';

    /// A quote '
    pub(crate) fn single_quote(input: CSpan<'_>) -> CNomResult<'_> {
        recognize(nchar(SINGLE_QUOTE))(input)
    }

    /// A string containing double ''' and ending (excluding) with a quote '
    pub(crate) fn string_esc_single_quote(input: CSpan<'_>) -> CNomResult<'_> {
        recognize(many0(alt((
            take_while1(|v| v != SINGLE_QUOTE),
            recognize(count(nchar(SINGLE_QUOTE), 2)),
        ))))(input)
    }

    /// SingleQuoted ::= "'" ([^'] | "''")+ "'"
    /// Parse a quoted string. A double quote within is an escaped quote.
    /// Returns the string within the outer quotes. The double quotes are not
    /// reduced.
    pub(crate) fn single_quoted_string(input: CSpan<'_>) -> CNomResult<'_> {
        recognize(tuple((
            error_code(single_quote, CRSingleQuoteStart),
            error_code(string_esc_single_quote, CRString),
            error_code(single_quote, CRSingleQuoteEnd),
        )))(input)
    }

    /// Hashtag
    pub(crate) fn hashtag(input: CSpan<'_>) -> CNomResult<'_> {
        error_code(tag("#"), CRHash)(input)
    }

    /// Sheet name
    pub(crate) fn sheet_name(input: CSpan<'_>) -> CParserResult<'_, CSpan<'_>> {
        let (rest, name) = match preceded(opt(dollar_nom), single_quoted_string)(input) {
            Ok((rest, name)) => (rest, name),
            Err(mut e) => match unquoted_sheet_name(input) {
                Ok((rest, name)) => (rest, name),
                Err(e2) => {
                    e.append(e2)?;
                    return Err(e);
                }
            },
        };

        Ok((rest, name))
    }

    /// SheetName ::= QuotedSheetName | '$'? [^\]\. #$']+
    /// QuotedSheetName ::= '$'? SingleQuoted
    pub(crate) fn unquoted_sheet_name(i: CSpan<'_>) -> CNomResult<'_> {
        recognize(error_code(many1(none_of(":]. #$'")), CRUnquotedName))(i)
    }

    /// Parse dollar
    pub(crate) fn dollar_nom(input: CSpan<'_>) -> CNomResult<'_> {
        error_code(tag("$"), CRDollar)(input)
    }

    /// Parse dot
    pub(crate) fn dot(input: CSpan<'_>) -> CNomResult<'_> {
        error_code(tag("."), CRDot)(input)
    }

    /// Parse colon
    pub(crate) fn colon(input: CSpan<'_>) -> CNomResult<'_> {
        error_code(tag(":"), CRColon)(input)
    }

    // Column ::= '$'? [A-Z]+
    /// Column label
    pub(crate) fn col(i: CSpan<'_>) -> CParserResult<'_, (Option<CSpan<'_>>, CSpan<'_>)> {
        let (i, abs) = opt(error_code(tag("$"), CRDollar))(i)?;
        let (i, col) = error_code(alpha1::<_, CParserError<'_>>, CRCol)(i)?;
        Ok((i, (abs, col)))
    }

    // Row ::= '$'? [1-9] [0-9]*
    /// Row label
    pub(crate) fn row(i: CSpan<'_>) -> CParserResult<'_, (Option<CSpan<'_>>, CSpan<'_>)> {
        let (i, abs) = opt(error_code(tag("$"), CRDollar))(i)?;
        let (i, row) = recognize(error_code(digit1, CRRow))(i)?;
        Ok((i, (abs, row)))
    }
}

#[cfg(test)]
mod tests {
    use crate::refs_impl::parser::tokens::{col, row};
    use crate::refs_impl::parser::CRCode::*;
    use crate::refs_impl::parser::{
        parse_cell_range, parse_cell_ref, parse_col_range, parse_row_range,
    };
    use crate::{CellRange, CellRef, ColRange, RowRange};
    use kparse::test::{str_parse, CheckTrace};
    use nom::error::ErrorKind;

    const R: CheckTrace = CheckTrace;

    #[test]
    pub(crate) fn test_col() {
        str_parse(&mut None, "", col).err(CRCol).q(R);
        str_parse(&mut None, "$A", col).ok_any().q(R);
        str_parse(&mut None, "$", col).err(CRCol).q(R);
        str_parse(&mut None, "A", col).ok_any().q(R);
        str_parse(&mut None, "$A ", col).ok_any().rest(" ").q(R);
    }

    #[test]
    pub(crate) fn test_row() {
        str_parse(&mut None, "", row).err(CRRow).q(R);
        str_parse(&mut None, "$1", row).ok_any().q(R);
        str_parse(&mut None, "$", row).err(CRRow).q(R);
        str_parse(&mut None, "1", row).ok_any().q(R);
        str_parse(&mut None, "$1 ", row).ok_any().rest(" ").q(R);
    }

    #[test]
    pub(crate) fn test_cellref() {
        fn iri(result: &CellRef, test: &str) -> bool {
            match result.iri() {
                Some(iri) => iri == test,
                None => false,
            }
        }
        fn table(result: &CellRef, test: &str) -> bool {
            match result.table() {
                Some(table) => table == test,
                None => false,
            }
        }
        fn row_col(result: &CellRef, test: &(u32, u32)) -> bool {
            (result.row(), result.col()) == *test
        }
        fn absolute(result: &CellRef, test: &(bool, bool)) -> bool {
            (result.row_abs(), result.col_abs()) == *test
        }

        str_parse(&mut None, "", parse_cell_ref).err_any().q(R);
        str_parse(&mut None, "'iri'#.A1", parse_cell_ref)
            .ok(iri, "iri")
            .q(R);
        str_parse(&mut None, "'iri'#.A1", parse_cell_ref)
            .ok(iri, "iri")
            .q(R);
        str_parse(&mut None, "'sheet'.A1", parse_cell_ref)
            .ok(table, "sheet")
            .q(R);
        str_parse(&mut None, ".A1", parse_cell_ref)
            .ok(row_col, &(0, 0))
            .ok(absolute, &(false, false))
            .q(R);
        str_parse(&mut None, ".A", parse_cell_ref).err(CRRow).q(R);
        str_parse(&mut None, ".1", parse_cell_ref).err(CRCol).q(R);
        str_parse(&mut None, "A1", parse_cell_ref).err(CRDot).q(R);
        str_parse(&mut None, ".$A$1", parse_cell_ref)
            .ok(row_col, &(0, 0))
            .ok(absolute, &(true, true))
            .q(R);
        str_parse(&mut None, ".$A $1", parse_cell_ref)
            .err(CRRow)
            .q(R);
        str_parse(&mut None, ".$ A$1", parse_cell_ref)
            .err(CRCol)
            .q(R);
        str_parse(&mut None, ".$A$ 1", parse_cell_ref)
            .err(CRRow)
            .q(R);
        str_parse(&mut None, ".$A$$1", parse_cell_ref)
            .err(CRRow)
            .q(R);
        str_parse(&mut None, ".$$A$$1", parse_cell_ref)
            .err(CRCol)
            .q(R);
        str_parse(&mut None, "'iri'#$'sheet'.$A$1", parse_cell_ref)
            .ok(iri, "iri")
            .ok(table, "sheet")
            .q(R);
    }

    #[test]
    pub(crate) fn test_cellrange() {
        fn iri(result: &CellRange, test: &str) -> bool {
            match result.iri() {
                Some(iri) => iri == test,
                None => false,
            }
        }
        fn table(result: &CellRange, test: &str) -> bool {
            match result.table() {
                Some(table) => table == test,
                None => false,
            }
        }
        fn to_table(result: &CellRange, test: &str) -> bool {
            match result.to_table() {
                Some(to_table) => to_table == test,
                None => false,
            }
        }
        fn row_col(result: &CellRange, test: &(u32, u32)) -> bool {
            (result.row(), result.col()) == *test
        }
        fn to_row_col(result: &CellRange, test: &(u32, u32)) -> bool {
            (result.to_row(), result.to_col()) == *test
        }
        fn absolute(result: &CellRange, test: &(bool, bool)) -> bool {
            (result.row_abs(), result.col_abs()) == *test
        }
        fn to_absolute(result: &CellRange, test: &(bool, bool)) -> bool {
            (result.to_row_abs(), result.to_col_abs()) == *test
        }

        str_parse(&mut None, "", parse_cell_range).err(CRDot).q(R);
        str_parse(&mut None, "'iri'#.A1:.C3", parse_cell_range)
            .ok(iri, "iri")
            .q(R);
        str_parse(&mut None, "'sheet'.A1:.C3", parse_cell_range)
            .ok(table, "sheet")
            .q(R);
        str_parse(&mut None, ".A1:.C3", parse_cell_range)
            .ok(row_col, &(0, 0))
            .ok(to_row_col, &(2, 2))
            .q(R);
        str_parse(&mut None, ".$A$1:.$C$3", parse_cell_range)
            .ok(row_col, &(0, 0))
            .ok(absolute, &(true, true))
            .ok(to_row_col, &(2, 2))
            .ok(to_absolute, &(true, true))
            .q(R);
        str_parse(&mut None, "'fun'.$A$1:'nofun'.$C$3", parse_cell_range)
            .ok(table, "fun")
            .ok(to_table, "nofun")
            .q(R);
        str_parse(&mut None, ".A1:.C3", parse_cell_range)
            .ok_any()
            .q(R);
        str_parse(&mut None, ".A1:.3", parse_cell_range)
            .err(CRCol)
            .nom_err(ErrorKind::Alpha)
            .expect(CRCol)
            .q(R);
        str_parse(&mut None, ".A1:.C", parse_cell_range)
            .err(CRRow)
            .nom_err(ErrorKind::Digit)
            .expect(CRRow)
            .q(R);
        str_parse(&mut None, ".A:.C3", parse_cell_range)
            .err(CRRow)
            .nom_err(ErrorKind::Digit)
            .expect(CRRow)
            .q(R);
        str_parse(&mut None, ".1:.C3", parse_cell_range)
            .err(CRCol)
            .nom_err(ErrorKind::Alpha)
            .expect(CRCol)
            .q(R);
        str_parse(&mut None, ":.C3", parse_cell_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, "A1:C3", parse_cell_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(
            &mut None,
            "'external'#'fun'.$A$1:'nofun'.$C$3",
            parse_cell_range,
        )
        .ok(table, "fun")
        .ok(to_table, "nofun")
        .q(R);
    }

    #[test]
    pub(crate) fn colrange() {
        fn iri(result: &ColRange, test: &str) -> bool {
            match result.iri() {
                Some(iri) => iri == test,
                None => false,
            }
        }
        fn sheet_name(result: &ColRange, test: &str) -> bool {
            match result.table() {
                Some(table) => table == test,
                None => false,
            }
        }
        fn col_col(result: &ColRange, test: &(u32, u32)) -> bool {
            (result.col(), result.to_col()) == *test
        }

        str_parse(&mut None, "", parse_col_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, "'iri'#.A:.C", parse_col_range)
            .ok(iri, "iri")
            .q(R);
        str_parse(&mut None, "'sheet'.A:.C", parse_col_range)
            .ok(sheet_name, "sheet")
            .q(R);
        str_parse(&mut None, ".A:.C", parse_col_range)
            .ok(col_col, &(0, 2))
            .q(R);
        str_parse(&mut None, ".1:", parse_col_range)
            .err(CRCol)
            .expect(CRCol)
            .nom_err(ErrorKind::Alpha)
            .q(R);
        str_parse(&mut None, ".A", parse_col_range)
            .err(CRColon)
            .expect(CRColon)
            .q(R);
        str_parse(&mut None, ":", parse_col_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, ":.A", parse_col_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, ":A", parse_col_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, ".5:.7", parse_col_range)
            .err(CRCol)
            .nom_err(ErrorKind::Alpha)
            .q(R);
        str_parse(&mut None, "'iri'#'sheet'.$A:.$C", parse_col_range)
            .ok(iri, "iri")
            .ok(sheet_name, "sheet")
            .q(R);
    }

    #[test]
    pub(crate) fn rowrange() {
        fn iri(result: &RowRange, test: &str) -> bool {
            match result.iri() {
                Some(iri) => iri == test,
                None => false,
            }
        }
        fn sheet_name(result: &RowRange, test: &str) -> bool {
            match result.table() {
                Some(table) => table == test,
                None => false,
            }
        }
        fn row_row(result: &RowRange, test: &(u32, u32)) -> bool {
            (result.row(), result.to_row()) == *test
        }

        str_parse(&mut None, "", parse_row_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, "'iri'#.1:.3", parse_row_range)
            .ok(iri, "iri")
            .q(R);
        str_parse(&mut None, "'sheet'.1:.3", parse_row_range)
            .ok(sheet_name, "sheet")
            .q(R);
        str_parse(&mut None, ".1:.3", parse_row_range)
            .ok(row_row, &(0, 2))
            .q(R);
        str_parse(&mut None, ".1:", parse_row_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, ".1", parse_row_range)
            .err(CRColon)
            .expect(CRColon)
            .q(R);
        str_parse(&mut None, ":", parse_row_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, ":.1", parse_row_range)
            .err(CRDot)
            .expect(CRDot)
            .q(R);
        str_parse(&mut None, ".C:.E", parse_row_range)
            .err(CRRow)
            .nom_err(ErrorKind::Digit)
            .q(R);
        str_parse(&mut None, "'iri'#'sheet'.$1:.$3", parse_row_range)
            .ok(iri, "iri")
            .ok(sheet_name, "sheet")
            .q(R);
    }
}
