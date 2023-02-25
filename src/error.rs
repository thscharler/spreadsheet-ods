//!
//! Error type.
//!

use kparse::tracker::TrackSpan;
use kparse::{Code, TokenizerError};
use std::error::Error;
use std::fmt::{Debug, Display, Formatter};
use std::str::from_utf8;

pub(crate) trait AsStatic<T: ?Sized> {
    fn as_static(&self) -> &'static T;
}

#[derive(Debug)]
#[allow(missing_docs)]
pub enum OdsError {
    Ods(String),
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    XmlAttr(quick_xml::events::attributes::AttrError),
    Utf8(std::str::Utf8Error),
    Parse(&'static str, Option<String>),
    ParseInt(std::num::ParseIntError),
    ParseBool(std::str::ParseBoolError),
    ParseFloat(std::num::ParseFloatError),
    Chrono(chrono::format::ParseError),
    SystemTime(std::time::SystemTimeError),
}

impl Display for OdsError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            OdsError::Ods(e) => write!(f, "Ods {}", e)?,
            OdsError::Io(e) => write!(f, "IO {}", e)?,
            OdsError::Zip(e) => write!(f, "Zip {:?}", e)?,
            OdsError::Xml(e) => write!(f, "Xml {}", e)?,
            OdsError::XmlAttr(e) => write!(f, "Xml attribute {}", e)?,
            OdsError::Parse(e, v) => write!(f, "Parse {} {:?}", e, v)?,
            OdsError::ParseInt(e) => write!(f, "ParseInt {}", e)?,
            OdsError::ParseBool(e) => write!(f, "ParseBool {}", e)?,
            OdsError::ParseFloat(e) => write!(f, "ParseFloat {}", e)?,
            OdsError::Chrono(e) => write!(f, "Chrono {}", e)?,
            OdsError::SystemTime(e) => write!(f, "SystemTime {}", e)?,
            OdsError::Utf8(e) => write!(f, "UTF8 {}", e)?,
        }

        Ok(())
    }
}

impl Error for OdsError {
    fn cause(&self) -> Option<&dyn Error> {
        match self {
            OdsError::Ods(_) => None,
            OdsError::Io(e) => Some(e),
            OdsError::Zip(e) => Some(e),
            OdsError::Xml(e) => Some(e),
            OdsError::XmlAttr(e) => Some(e),
            OdsError::Parse(_, _) => None,
            OdsError::ParseInt(e) => Some(e),
            OdsError::ParseBool(e) => Some(e),
            OdsError::ParseFloat(e) => Some(e),
            OdsError::Chrono(e) => Some(e),
            OdsError::SystemTime(e) => Some(e),
            OdsError::Utf8(e) => Some(e),
        }
    }
}

impl From<std::io::Error> for OdsError {
    fn from(err: std::io::Error) -> OdsError {
        OdsError::Io(err)
    }
}

impl From<zip::result::ZipError> for OdsError {
    fn from(err: zip::result::ZipError) -> OdsError {
        OdsError::Zip(err)
    }
}

impl From<quick_xml::Error> for OdsError {
    fn from(err: quick_xml::Error) -> OdsError {
        OdsError::Xml(err)
    }
}

impl From<quick_xml::events::attributes::AttrError> for OdsError {
    fn from(err: quick_xml::events::attributes::AttrError) -> OdsError {
        OdsError::XmlAttr(err)
    }
}

impl From<std::str::ParseBoolError> for OdsError {
    fn from(err: std::str::ParseBoolError) -> OdsError {
        OdsError::ParseBool(err)
    }
}

impl From<std::num::ParseIntError> for OdsError {
    fn from(err: std::num::ParseIntError) -> OdsError {
        OdsError::ParseInt(err)
    }
}

impl From<std::num::ParseFloatError> for OdsError {
    fn from(err: std::num::ParseFloatError) -> OdsError {
        OdsError::ParseFloat(err)
    }
}

impl From<chrono::format::ParseError> for OdsError {
    fn from(err: chrono::format::ParseError) -> OdsError {
        OdsError::Chrono(err)
    }
}

impl From<std::time::SystemTimeError> for OdsError {
    fn from(err: std::time::SystemTimeError) -> OdsError {
        OdsError::SystemTime(err)
    }
}

impl From<std::str::Utf8Error> for OdsError {
    fn from(err: std::str::Utf8Error) -> OdsError {
        OdsError::Utf8(err)
    }
}

impl<C> From<nom::Err<TokenizerError<C, &[u8]>>> for OdsError
where
    C: AsStatic<str>,
{
    fn from(value: nom::Err<TokenizerError<C, &[u8]>>) -> Self {
        match value {
            nom::Err::Incomplete(_) => OdsError::Parse("incomplete", None),
            nom::Err::Error(e) | nom::Err::Failure(e) => OdsError::Parse(
                e.code.as_static(),
                Some(from_utf8(e.span).unwrap_or("decoding failed").into()),
            ),
        }
    }
}

impl<C> From<nom::Err<TokenizerError<C, &str>>> for OdsError
where
    C: AsStatic<str>,
{
    fn from(value: nom::Err<TokenizerError<C, &str>>) -> Self {
        match value {
            nom::Err::Incomplete(_) => OdsError::Parse("incomplete", None),
            nom::Err::Error(e) | nom::Err::Failure(e) => {
                OdsError::Parse(e.code.as_static(), Some(e.span.into()))
            }
        }
    }
}

impl<'s, C> From<nom::Err<TokenizerError<C, TrackSpan<'s, C, &'s str>>>> for OdsError
where
    C: Code + AsStatic<str>,
{
    fn from(value: nom::Err<TokenizerError<C, TrackSpan<'s, C, &'s str>>>) -> Self {
        match value {
            nom::Err::Incomplete(_) => OdsError::Parse("incomplete", None),
            nom::Err::Error(e) | nom::Err::Failure(e) => {
                OdsError::Parse(e.code.as_static(), Some((*e.span).into()))
            }
        }
    }
}
