// !!
// !! clean copy from openformula crate.
// !! do not modify, except for use clauses.
// !!

//!
//! Low level conversions from a Span<'a> to ...
//!

use crate::refs_impl::Span;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::num::IntErrorKind;
use std::str::FromStr;

/// Replaces two quotes (") with a single one.
/// Strips one leading and one trailing quote.
pub fn unquote_double(i: Span<'_>) -> String {
    let i = match i.strip_prefix('\"') {
        None => *i,
        Some(s) => s,
    };
    let i = match i.strip_suffix('\"') {
        None => i,
        Some(s) => s,
    };

    (*i).replace(r#""""#, r#"""#)
}

/// Replaces one quote (") with two.
pub fn quote_double(i: &str) -> String {
    (*i).replace('"', "\"\"")
}

/// Replaces two single quotes (') with a single on.
/// Strips one leading and one trailing quote.
pub fn unquote_single(i: Span<'_>) -> String {
    let i = match i.strip_prefix('\'') {
        None => *i,
        Some(s) => s,
    };
    let i = match i.strip_suffix('\'') {
        None => i,
        Some(s) => s,
    };

    (*i).replace("''", "'")
}

/// Replaces one single quote (') with two.
pub fn quote_single(i: &str) -> String {
    (*i).replace('\'', "''")
}

/// Parse a bool if a '$' exists.
pub fn try_bool_from_abs_flag(i: Option<Span<'_>>) -> bool {
    if let Some(i) = i {
        *i == "$"
    } else {
        false
    }
}

/// Error for try_u32_from_rowname.
#[allow(variant_size_differences)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ParseRownameError {
    /// Value being parsed is empty.
    ///
    /// This variant will be constructed when parsing an empty string.
    Empty,
    /// Contains an invalid digit in its context.
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
pub fn try_u32_from_rowname(i: Span<'_>) -> Result<u32, ParseRownameError> {
    match u32::from_str(*i) {
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
pub enum ParseColnameError {
    /// Invalid column character.
    InvalidChar(char),
    /// Invalid column name.
    InvalidColname(String),
}

impl Display for ParseColnameError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            ParseColnameError::InvalidChar(e) => {
                write!(f, "Invalid char '{}'", e)?;
            }
            ParseColnameError::InvalidColname(e) => {
                write!(f, "Invalid colname {}", e)?;
            }
        }
        Ok(())
    }
}

impl Error for ParseColnameError {}

/// Parse a col label to a column index.
pub fn try_u32_from_colname(i: Span<'_>) -> Result<u32, ParseColnameError> {
    let mut col = 0u32;

    for c in (*i).chars() {
        if !('A'..='Z').contains(&c) {
            return Err(ParseColnameError::InvalidChar(c));
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
        Err(ParseColnameError::InvalidColname(format!("{:?}", i)))
    } else {
        Ok(col - 1)
    }
}
