//!
//! Parses different data from a &[u8].
//!
//! For many cases this omits the transformation to a &str

use crate::{OdsError, Visibility};
use chrono::Duration;
use chrono::{NaiveDateTime, ParseResult};
use nom::branch::alt;
use nom::bytes::complete::tag;
use nom::character::complete::digit1;
use nom::combinator::{eof, map, opt, recognize};
use nom::error::{ErrorKind, FromExternalError};
use nom::number::complete::double;
use nom::sequence::{pair, preceded, terminated, tuple};
use nom::{IResult, Slice};
use quick_xml::escape::unescape;
use std::borrow::Cow;
use std::str::{from_utf8, from_utf8_unchecked};

/// Unescape and decode as UTF8
pub(crate) fn parse_string(input: &Cow<'_, [u8]>) -> Result<String, OdsError> {
    let result = match unescape(input.as_ref()) {
        Ok(result) => result,
        Err(err) => return Err(OdsError::Parse(err.to_string())),
    };

    let result = from_utf8(result.as_ref())?;

    Ok(result.to_string())
}

/// Parse as Visibility.
pub(crate) fn parse_visibility(input: &Cow<'_, [u8]>) -> Result<Visibility, OdsError> {
    match input.as_ref() {
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
pub(crate) fn parse_currency(input: &Cow<'_, [u8]>) -> Result<[u8; 3], OdsError> {
    let v = input.as_ref();
    let result = match v.len() {
        0 => Ok([b' ', b' ', b' ']),
        1 => Ok([v[0], b' ', b' ']),
        2 => Ok([v[0], v[1], b' ']),
        3 => Ok([v[0], v[1], v[2]]),
        _ => Err(OdsError::Parse(format!("{:?} not a currency", v))),
    };

    result
}

/// Parse a bool.
pub(crate) fn parse_bool(input: &Cow<'_, [u8]>) -> Result<bool, OdsError> {
    Ok(token_bool(input.as_ref())?.1)
}

/// Parse a u32.
pub(crate) fn parse_u32(input: &Cow<'_, [u8]>) -> Result<u32, OdsError> {
    Ok(token_u32(input.as_ref())?.1)
}

/// Parse a i64.
pub(crate) fn parse_i64(input: &Cow<'_, [u8]>) -> Result<i64, OdsError> {
    Ok(token_i64(input.as_ref())?.1)
}

/// Parse a i32.
pub(crate) fn parse_i32(input: &Cow<'_, [u8]>) -> Result<i32, OdsError> {
    Ok(token_i32(input.as_ref())?.1)
}

/// Parse a i16.
pub(crate) fn parse_i16(input: &Cow<'_, [u8]>) -> Result<i16, OdsError> {
    Ok(token_i16(input.as_ref())?.1)
}

/// Parse a f64.
pub(crate) fn parse_f64(input: &Cow<'_, [u8]>) -> Result<f64, OdsError> {
    Ok(token_float(input.as_ref())?.1)
}

/// Parse a XML Schema datetime.
pub(crate) fn parse_datetime(input: &Cow<'_, [u8]>) -> Result<NaiveDateTime, OdsError> {
    Ok(token_datetime(input.as_ref())?.1)
}

/// Parse a XML Schema time duration.
pub(crate) fn parse_duration(input: &Cow<'_, [u8]>) -> Result<Duration, OdsError> {
    Ok(token_duration(input.as_ref())?.1)
}

fn token_bool(input: &[u8]) -> IResult<&[u8], bool> {
    let (input, result) = terminated(
        alt((map(tag(b"true"), |_| true), map(tag(b"false"), |_| false))),
        eof,
    )(input)?;
    Ok((input, result))
}

fn token_i16(input: &[u8]) -> IResult<&[u8], i16> {
    let (input, result) = terminated(recognize(tuple((opt(byte(b'-')), digit1))), eof)(input)?;

    let result = match unsafe { from_utf8_unchecked(result) }.parse::<i16>() {
        Ok(result) => result,
        Err(err) => {
            return Err(nom::Err::Error(nom::error::Error::from_external_error(
                input,
                ErrorKind::Verify,
                err,
            )))
        }
    };

    Ok((input, result))
}

fn token_u32(input: &[u8]) -> IResult<&[u8], u32> {
    let (input, result) = terminated(digit1, eof)(input)?;

    let result = match unsafe { from_utf8_unchecked(result) }.parse::<u32>() {
        Ok(result) => result,
        Err(err) => {
            return Err(nom::Err::Error(nom::error::Error::from_external_error(
                input,
                ErrorKind::Verify,
                err,
            )))
        }
    };

    Ok((input, result))
}

fn token_i32(input: &[u8]) -> IResult<&[u8], i32> {
    let (input, result) = terminated(recognize(tuple((opt(byte(b'-')), digit1))), eof)(input)?;

    let result = match unsafe { from_utf8_unchecked(result) }.parse::<i32>() {
        Ok(result) => result,
        Err(err) => {
            return Err(nom::Err::Error(nom::error::Error::from_external_error(
                input,
                ErrorKind::Verify,
                err,
            )))
        }
    };

    Ok((input, result))
}

fn token_i64(input: &[u8]) -> IResult<&[u8], i64> {
    let (input, result) = terminated(recognize(tuple((opt(byte(b'-')), digit1))), eof)(input)?;

    let result = match unsafe { from_utf8_unchecked(result) }.parse::<i64>() {
        Ok(result) => result,
        Err(err) => {
            return Err(nom::Err::Error(nom::error::Error::from_external_error(
                input,
                ErrorKind::Verify,
                err,
            )))
        }
    };

    Ok((input, result))
}

fn token_float(input: &[u8]) -> IResult<&[u8], f64> {
    let (input, result) = terminated(double, eof)(input)?;

    Ok((input, result))
}

// Part of a date/duration. An unsigned integer, but for chrono we need an i64.
fn token_datepart(input: &[u8]) -> IResult<&[u8], i64> {
    let (input, result) = digit1(input)?;

    let result = match unsafe { from_utf8_unchecked(result) }.parse::<i64>() {
        Ok(result) => result,
        Err(err) => {
            return Err(nom::Err::Error(nom::error::Error::from_external_error(
                input,
                ErrorKind::Verify,
                err,
            )))
        }
    };

    Ok((input, result))
}

// Part of a date/duration. Parses an integer as nanoseconds but with
// the caveat that there can be arbitrary many 0s omitted.
fn token_nano(input: &[u8]) -> IResult<&[u8], i64> {
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

fn to_naive_datetime(
    fields: (i64, i64, i64, Option<(i64, i64, i64, Option<i64>)>),
) -> ParseResult<NaiveDateTime> {
    let mut p = chrono::format::Parsed::new();
    p.set_year(fields.0)?;
    p.set_month(fields.1)?;
    p.set_day(fields.2)?;
    if let Some(fields) = fields.3 {
        p.set_hour(fields.0)?;
        p.set_minute(fields.1)?;
        p.set_second(fields.2)?;
        if let Some(fields) = fields.3 {
            p.set_nanosecond(fields)?;
        }
    } else {
        p.set_hour(0)?;
        p.set_minute(0)?;
        p.set_second(0)?;
    }
    p.to_naive_datetime_with_offset(0)
}

fn token_datetime(input: &[u8]) -> IResult<&[u8], NaiveDateTime> {
    let (input, result) = terminated(
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

    let sign = match result.0 {
        Some(_) => -1,
        None => 1,
    };

    let result = match to_naive_datetime((
        sign * result.1,
        result.3,
        result.5,
        if let Some(result) = result.6 {
            Some((
                result.1,
                result.3,
                result.5,
                if let Some(result) = result.6 {
                    Some(result.1)
                } else {
                    None
                },
            ))
        } else {
            None
        },
    )) {
        Ok(result) => result,
        Err(err) => {
            return Err(nom::Err::Error(nom::error::Error::from_external_error(
                input,
                ErrorKind::Verify,
                err,
            )));
        }
    };

    Ok((input, result))
}

fn token_duration(input: &[u8]) -> IResult<&[u8], Duration> {
    let (input, result) = terminated(
        tuple((
            byte(b'P'),
            opt(terminated(token_datepart, byte(b'Y'))),
            opt(terminated(token_datepart, byte(b'M'))),
            opt(terminated(token_datepart, byte(b'D'))),
            byte(b'T'),
            terminated(token_datepart, byte(b'H')),
            terminated(token_datepart, byte(b'M')),
            terminated(
                pair(token_datepart, opt(preceded(byte(b'.'), token_nano))),
                byte(b'S'),
            ),
        )),
        eof,
    )(input)?;

    let result = if let Some(nanos) = result.7 .1 {
        Duration::seconds(result.5 * 3600 + result.6 * 60 + result.7 .0)
            + Duration::nanoseconds(nanos)
    } else {
        Duration::seconds(result.5 * 3600 + result.6 * 60 + result.7 .0)
    };

    Ok((input, result))
}

pub(crate) fn byte(c: u8) -> impl Fn(&[u8]) -> IResult<&[u8], u8> {
    move |i: &[u8]| match i.iter().next() {
        Some(x) if *x == c => Ok((i.slice(1..), *x)),
        _ => Err(nom::Err::Error(nom::error::Error::new(i, ErrorKind::Char))),
    }
}

#[cfg(test)]
mod tests {
    use crate::io::parse::{
        parse_bool, parse_datetime, parse_duration, parse_f64, parse_i32, parse_string, parse_u32,
        token_nano,
    };
    use crate::OdsError;
    use std::borrow::Cow;

    #[test]
    fn test_string() -> Result<(), OdsError> {
        assert_eq!(parse_string(&Cow::Borrowed(b"a&lt;sdf"))?, "a<sdf");
        assert_eq!(parse_string(&Cow::Borrowed(b"asdf"))?, "asdf");

        Ok(())
    }

    #[test]
    fn test_u32() -> Result<(), OdsError> {
        assert_eq!(parse_u32(&Cow::Borrowed(b"1234"))?, 1234);
        parse_u32(&Cow::Borrowed(b"123456789000")).unwrap_err();
        parse_u32(&Cow::Borrowed(b"1234 ")).unwrap_err();
        parse_u32(&Cow::Borrowed(b"-1234 ")).unwrap_err();
        parse_u32(&Cow::Borrowed(b"-1234")).unwrap_err();

        Ok(())
    }

    #[test]
    fn test_i32() -> Result<(), OdsError> {
        assert_eq!(parse_i32(&Cow::Borrowed(b"1234"))?, 1234);
        assert_eq!(parse_i32(&Cow::Borrowed(b"-1234"))?, -1234);
        parse_i32(&Cow::Borrowed(b"1234 ")).unwrap_err();
        parse_i32(&Cow::Borrowed(b"-1234 ")).unwrap_err();
        parse_i32(&Cow::Borrowed(b"123456789000")).unwrap_err();

        Ok(())
    }

    #[test]
    fn test_float() -> Result<(), OdsError> {
        assert_eq!(parse_f64(&Cow::Borrowed(b"1234"))?, 1234.);
        assert_eq!(parse_f64(&Cow::Borrowed(b"-1234"))?, -1234.);
        assert_eq!(parse_f64(&Cow::Borrowed(b"123456789000"))?, 123456789000.);
        assert_eq!(parse_f64(&Cow::Borrowed(b"1234.5678"))?, 1234.5678);
        parse_f64(&Cow::Borrowed(b"1234 ")).unwrap_err();
        parse_f64(&Cow::Borrowed(b"-1234 ")).unwrap_err();

        Ok(())
    }

    #[test]
    fn test_datetime() -> Result<(), OdsError> {
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"19999-01-01"))?.timestamp(),
            568940284800
        );
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"1999-01-01"))?.timestamp(),
            915148800
        );
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"-45-01-01"))?.timestamp(),
            -63587289600
        );
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"2004-02-29"))?.timestamp(),
            1078012800
        );
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"2000-02-29"))?.timestamp(),
            951782400
        );

        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"2000-01-01T11:22:33"))?.timestamp(),
            946725753
        );
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"2000-01-01T11:22:33.1234"))?.timestamp(),
            946725753
        );
        assert_eq!(
            parse_datetime(&Cow::Borrowed(b"2000-01-01T11:22:33.123456789111"))?.timestamp(),
            946725753
        );

        Ok(())
    }

    #[test]
    fn test_duration() -> Result<(), OdsError> {
        assert_eq!(
            parse_duration(&Cow::Borrowed(b"PT12H12M12S"))?.num_milliseconds(),
            43932000
        );
        assert_eq!(
            parse_duration(&Cow::Borrowed(b"PT12H12M12.223S"))?.num_milliseconds(),
            43932223
        );
        Ok(())
    }

    #[test]
    fn test_bool() -> Result<(), OdsError> {
        assert_eq!(parse_bool(&Cow::Borrowed(b"true"))?, true);
        assert_eq!(parse_bool(&Cow::Borrowed(b"false"))?, false);
        parse_bool(&Cow::Borrowed(b"ffoso")).unwrap_err();
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
