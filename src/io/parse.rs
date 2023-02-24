//!
//! Parses different data from a &[u8].
//!
//! For many cases this omits the transformation to a &str

use crate::error::AsStatic;
use crate::{OdsError, Visibility};
use chrono::Duration;
use chrono::NaiveDateTime;
use kparse::{Code, KParser, TokenizerError, TokenizerResult};
use nom::bytes::complete::tag;
use nom::character::complete::digit1;
use nom::combinator::{all_consuming, eof, opt};
use nom::number::complete::double;
use nom::sequence::{pair, preceded, terminated, tuple};
use nom::{AsChar, Parser};
use std::fmt::{Display, Formatter};
use std::str::{from_utf8, from_utf8_unchecked};

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub(crate) enum RCode {
    NomError,

    Byte,
    Digit,

    Integer,
    DateTime,
}

impl AsStatic<str> for RCode {
    fn as_static(&self) -> &'static str {
        match self {
            RCode::NomError => "NomError",
            RCode::Byte => "Byte",
            RCode::Digit => "Digit",
            RCode::Integer => "Integer",
            RCode::DateTime => "DateTime",
        }
    }
}

impl Display for RCode {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{:?}", self)
    }
}

impl Code for RCode {
    const NOM_ERROR: Self = Self::NomError;
}

// #[cfg(debug_assertions)]
// pub(crate) type KSpan<'s> = TrackSpan<'s, ACode, &'s [u8]>;
// #[cfg(not(debug_assertions))]
pub(crate) type KSpan<'s> = &'s [u8];
// pub(crate) type KParserResult<'s, O> = ParserResult<ACode, KSpan<'s>, O>;
pub(crate) type KTokenizerResult<'s, O> = TokenizerResult<RCode, KSpan<'s>, O>;
// pub(crate) type KParserError<'s> = ParserError<ACode, KSpan<'s>>;
pub(crate) type KTokenizerError<'s> = TokenizerError<RCode, KSpan<'s>>;

/// Parse as Visibility.
pub(crate) fn parse_visibility(input: &[u8]) -> Result<Visibility, OdsError> {
    match input {
        b"visible" => Ok(Visibility::Visible),
        b"filter" => Ok(Visibility::Filtered),
        b"collapse" => Ok(Visibility::Collapsed),
        _ => Err(OdsError::Ods(format!(
            "Unknown value for table:visibility {}",
            from_utf8(input)?
        ))),
    }
}

/// Parse a attribute value as a currency.
pub(crate) fn parse_currency(input: &[u8]) -> Result<[u8; 3], OdsError> {
    match input.len() {
        0 => Ok([b' ', b' ', b' ']),
        1 => Ok([input[0], b' ', b' ']),
        2 => Ok([input[0], input[1], b' ']),
        3 => Ok([input[0], input[1], input[2]]),
        _ => Err(OdsError::Parse(
            "not a currency",
            Some(from_utf8(input)?.into()),
        )),
    }
}

/// Parse a bool.
pub(crate) fn parse_bool(input: &[u8]) -> Result<bool, OdsError> {
    Ok(token_bool(input)?.1)
}

/// Parse a u32.
pub(crate) fn parse_u32(input: &[u8]) -> Result<u32, OdsError> {
    Ok(token_u32(input)?.1)
}

/// Parse a i64.
pub(crate) fn parse_i64(input: &[u8]) -> Result<i64, OdsError> {
    Ok(token_i64(input)?.1)
}

/// Parse a i32.
pub(crate) fn parse_i32(input: &[u8]) -> Result<i32, OdsError> {
    Ok(token_i32(input)?.1)
}

/// Parse a i16.
pub(crate) fn parse_i16(input: &[u8]) -> Result<i16, OdsError> {
    Ok(token_i16(input)?.1)
}

/// Parse a f64.
pub(crate) fn parse_f64(input: &[u8]) -> Result<f64, OdsError> {
    Ok(token_float(input)?.1)
}

/// Parse a XML Schema datetime.
pub(crate) fn parse_datetime(input: &[u8]) -> Result<NaiveDateTime, OdsError> {
    Ok(token_datetime(input)?.1)
}

/// Parse a XML Schema time duration.
pub(crate) fn parse_duration(input: &[u8]) -> Result<Duration, OdsError> {
    Ok(token_duration(input)?.1)
}

#[inline(always)]
fn token_bool(input: &[u8]) -> KTokenizerResult<'_, bool> {
    let (_rest, val) = tag::<_, _, TokenizerError<RCode, _>>(b"true")
        .value(true)
        .or(tag(b"false").value(false))
        .parse(input)?;
    Ok((input, val))
}

#[inline(always)]
fn token_i16(input: &[u8]) -> KTokenizerResult<'_, i16> {
    let _ = opt(byte(b'-')).and(all_digits).parse(input)?;

    let result = match unsafe { from_utf8_unchecked(input) }.parse::<i16>() {
        Ok(result) => result,
        Err(_) => {
            return Err(nom::Err::Error(KTokenizerError::new(RCode::Integer, input)));
        }
    };

    Ok((input, result))
}

#[inline(always)]
fn token_u32(input: &[u8]) -> KTokenizerResult<'_, u32> {
    let _ = all_digits.parse(input)?;

    let result = match unsafe { from_utf8_unchecked(input) }.parse::<u32>() {
        Ok(result) => result,
        Err(_) => {
            return Err(nom::Err::Error(KTokenizerError::new(RCode::Integer, input)));
        }
    };

    Ok((input, result))
}

#[inline(always)]
fn token_i32(input: &[u8]) -> KTokenizerResult<'_, i32> {
    let _ = opt(byte(b'-')).and(all_digits).parse(input)?;

    let result = match unsafe { from_utf8_unchecked(input) }.parse::<i32>() {
        Ok(result) => result,
        Err(_) => {
            return Err(nom::Err::Error(KTokenizerError::new(RCode::Integer, input)));
        }
    };

    Ok((input, result))
}

#[inline(always)]
fn token_i64(input: &[u8]) -> KTokenizerResult<'_, i64> {
    let _ = opt(byte(b'-')).and(all_digits).parse(input)?;

    let result = match unsafe { from_utf8_unchecked(input) }.parse::<i64>() {
        Ok(result) => result,
        Err(_) => {
            return Err(nom::Err::Error(KTokenizerError::new(RCode::Integer, input)));
        }
    };

    Ok((input, result))
}

#[inline(always)]
fn token_float(input: &[u8]) -> KTokenizerResult<'_, f64> {
    let (input, result) = all_consuming(double)(input)?;

    Ok((input, result))
}

// Part of a date/duration. An unsigned integer, but for chrono we need an i64.
#[inline]
fn token_datepart(input: &[u8]) -> KTokenizerResult<'_, i64> {
    let (input, result) = digit1(input)?;

    let result = match unsafe { from_utf8_unchecked(result) }.parse::<i64>() {
        Ok(result) => result,
        Err(_) => {
            return Err(nom::Err::Error(KTokenizerError::new(RCode::Integer, input)));
        }
    };

    Ok((input, result))
}

// Part of a date/duration. Parses an integer as nanoseconds but with
// the caveat that there can be arbitrary many trailing 0s omitted.
#[inline]
fn token_nano(input: &[u8]) -> KTokenizerResult<'_, i64> {
    let (input, result) = digit1(input)?;

    let mut v = 0i64;
    for i in 0..9 {
        if i < result.len() {
            v *= 10;
            v += (result[i] - b'0') as i64;
        } else {
            v *= 10;
        }
    }
    Ok((input, v))
}

#[inline(always)]
fn token_datetime(input: &[u8]) -> KTokenizerResult<'_, NaiveDateTime> {
    let (input, (minus, year, _, month, _, day, time)) = terminated(
        tuple((
            opt(byte(b'-')),
            token_datepart,
            byte(b'-'),
            token_datepart,
            byte(b'-'),
            token_datepart,
            opt(tuple((
                byte(b'T'),
                token_datepart,
                byte(b':'),
                token_datepart,
                byte(b':'),
                token_datepart,
                opt(tuple((byte(b'.'), token_nano))),
            ))),
        )),
        eof,
    )(input)?;

    let sign = match minus {
        Some(_) => -1,
        None => 1,
    };

    let mut p = chrono::format::Parsed::new();
    p.year = Some((sign * year) as i32);
    p.month = Some(month as u32);
    p.day = Some(day as u32);
    if let Some((_, hour, _, minute, _, second, nanos)) = time {
        p.hour_div_12 = Some((hour / 12) as u32);
        p.hour_mod_12 = Some((hour % 12) as u32);
        p.minute = Some(minute as u32);
        p.second = Some(second as u32);
        if let Some((_, nanos)) = nanos {
            p.nanosecond = Some(nanos as u32);
        }
    } else {
        p.hour_div_12 = Some(0);
        p.hour_mod_12 = Some(0);
        p.minute = Some(0);
        p.second = Some(0);
    }
    match p.to_naive_datetime_with_offset(0) {
        Ok(v) => Ok((input, v)),
        Err(err) => {
            dbg!(&err);
            return Err(nom::Err::Error(KTokenizerError::new(
                RCode::DateTime,
                input,
            )));
        }
    }
}

#[inline(always)]
fn token_duration(input: &[u8]) -> KTokenizerResult<'_, Duration> {
    let (input, (_, day, _, hour, minute, (second, nanos))) = all_consuming(tuple((
        byte(b'P'),
        // these do not occur?
        //opt(terminated(token_datepart, byte(b'Y'))),
        //opt(terminated(token_datepart, byte(b'M'))),
        opt(terminated(token_datepart, byte(b'D'))),
        byte(b'T'),
        terminated(token_datepart, byte(b'H')),
        terminated(token_datepart, byte(b'M')),
        terminated(
            pair(token_datepart, opt(preceded(byte(b'.'), token_nano))),
            byte(b'S'),
        ),
    )))(input)?;

    let mut result = Duration::seconds(hour * 3600 + minute * 60 + second);
    if let Some(day) = day {
        result = result + Duration::days(day);
    }
    if let Some(nanos) = nanos {
        result = result + Duration::nanoseconds(nanos);
    }

    Ok((input, result))
}

#[inline(always)]
fn all_digits(input: &[u8]) -> KTokenizerResult<'_, ()> {
    for c in input {
        if !c.is_dec_digit() {
            return Err(nom::Err::Error(KTokenizerError::new(RCode::Digit, input)));
        }
    }
    Ok((&input[input.len()..], ()))
}

#[inline(always)]
pub(crate) fn byte(c: u8) -> impl Fn(&[u8]) -> KTokenizerResult<'_, ()> {
    move |i: &[u8]| {
        if i.len() > 0 && i[0] == c {
            Ok((&i[1..], ()))
        } else {
            Err(nom::Err::Error(KTokenizerError::new(RCode::Byte, i)))
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::io::parse::{
        parse_bool, parse_datetime, parse_duration, parse_f64, parse_i32, parse_u32, token_nano,
    };
    use crate::OdsError;

    #[test]
    fn test_u32() -> Result<(), OdsError> {
        assert_eq!(parse_u32(b"1234")?, 1234);
        parse_u32(b"123456789000").unwrap_err();
        parse_u32(b"1234 ").unwrap_err();
        parse_u32(b"-1234 ").unwrap_err();
        parse_u32(b"-1234").unwrap_err();

        Ok(())
    }

    #[test]
    fn test_i32() -> Result<(), OdsError> {
        assert_eq!(parse_i32(b"1234")?, 1234);
        assert_eq!(parse_i32(b"-1234")?, -1234);
        parse_i32(b"1234 ").unwrap_err();
        parse_i32(b"-1234 ").unwrap_err();
        parse_i32(b"123456789000").unwrap_err();

        Ok(())
    }

    #[test]
    fn test_float() -> Result<(), OdsError> {
        assert_eq!(parse_f64(b"1234")?, 1234.);
        assert_eq!(parse_f64(b"-1234")?, -1234.);
        assert_eq!(parse_f64(b"123456789000")?, 123456789000.);
        assert_eq!(parse_f64(b"1234.5678")?, 1234.5678);
        parse_f64(b"1234 ").unwrap_err();
        parse_f64(b"-1234 ").unwrap_err();

        Ok(())
    }

    #[test]
    fn test_datetime() -> Result<(), OdsError> {
        assert_eq!(parse_datetime(b"19999-01-01")?.timestamp(), 568940284800);
        assert_eq!(parse_datetime(b"1999-01-01")?.timestamp(), 915148800);
        assert_eq!(parse_datetime(b"-45-01-01")?.timestamp(), -63587289600);
        assert_eq!(parse_datetime(b"2004-02-29")?.timestamp(), 1078012800);
        assert_eq!(parse_datetime(b"2000-02-29")?.timestamp(), 951782400);

        assert_eq!(
            parse_datetime(b"2000-01-01T11:22:33")?.timestamp(),
            946725753
        );
        assert_eq!(
            parse_datetime(b"2000-01-01T11:22:33.1234")?.timestamp(),
            946725753
        );
        assert_eq!(
            parse_datetime(b"2000-01-01T11:22:33.123456789111")?.timestamp(),
            946725753
        );

        Ok(())
    }

    #[test]
    fn test_duration() -> Result<(), OdsError> {
        assert_eq!(parse_duration(b"PT12H12M12S")?.num_milliseconds(), 43932000);
        assert_eq!(
            parse_duration(b"PT12H12M12.223S")?.num_milliseconds(),
            43932223
        );
        Ok(())
    }

    #[test]
    fn test_bool() -> Result<(), OdsError> {
        assert_eq!(parse_bool(b"true")?, true);
        assert_eq!(parse_bool(b"false")?, false);
        parse_bool(b"ffoso").unwrap_err();
        Ok(())
    }

    #[test]
    fn test_nano() -> Result<(), OdsError> {
        assert_eq!(token_nano(b"123")?.1, 123000000i64);
        assert_eq!(token_nano(b"123456789")?.1, 123456789i64);
        assert_eq!(token_nano(b"1234567897777")?.1, 123456789i64);
        Ok(())
    }
}
