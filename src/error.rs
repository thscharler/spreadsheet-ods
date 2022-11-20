//!
//! Error type.
//!

use crate::refs_impl::error::ParseOFError;
use spreadsheet_ods_cellref::CellRefError;
use std::fmt::{Display, Formatter};

#[derive(Debug)]
#[allow(missing_docs)]
pub enum OdsError {
    Ods(String),
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    XmlAttr(quick_xml::events::attributes::AttrError),
    Escape(String),
    Utf8(std::str::Utf8Error),
    Parse(String),
    ParseInt(std::num::ParseIntError),
    ParseBool(std::str::ParseBoolError),
    ParseFloat(std::num::ParseFloatError),
    Chrono(chrono::format::ParseError),
    SystemTime(std::time::SystemTimeError),
    Nom(nom::error::Error<String>),
    CellRef(String),
}

impl Display for OdsError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            OdsError::Ods(e) => write!(f, "Ods {}", e)?,
            OdsError::Io(e) => write!(f, "IO {}", e)?,
            OdsError::Zip(e) => write!(f, "Zip {:?}", e)?,
            OdsError::Xml(e) => write!(f, "Xml {}", e)?,
            OdsError::XmlAttr(e) => write!(f, "Xml attribute {}", e)?,
            OdsError::Parse(e) => write!(f, "Parse {}", e)?,
            OdsError::ParseInt(e) => write!(f, "ParseInt {}", e)?,
            OdsError::ParseBool(e) => write!(f, "ParseBool {}", e)?,
            OdsError::ParseFloat(e) => write!(f, "ParseFloat {}", e)?,
            OdsError::Chrono(e) => write!(f, "Chrono {}", e)?,
            OdsError::SystemTime(e) => write!(f, "SystemTime {}", e)?,
            OdsError::Utf8(e) => write!(f, "UTF8 {}", e)?,
            OdsError::Nom(e) => write!(f, "Nom {}", e)?,
            OdsError::Escape(s) => write!(f, "Escape {}", s)?,
            OdsError::CellRef(e) => write!(f, "CellRef {:?}", e)?,
        }

        Ok(())
    }
}

impl std::error::Error for OdsError {
    fn cause(&self) -> Option<&dyn std::error::Error> {
        match self {
            OdsError::Ods(_) => None,
            OdsError::Io(e) => Some(e),
            OdsError::Zip(e) => Some(e),
            OdsError::Xml(e) => Some(e),
            OdsError::XmlAttr(e) => Some(e),
            OdsError::Parse(_) => None,
            OdsError::ParseInt(e) => Some(e),
            OdsError::ParseBool(e) => Some(e),
            OdsError::ParseFloat(e) => Some(e),
            OdsError::Chrono(e) => Some(e),
            OdsError::SystemTime(e) => Some(e),
            OdsError::Utf8(e) => Some(e),
            OdsError::Nom(e) => Some(e),
            OdsError::Escape(_) => None,
            OdsError::CellRef(_) => None,
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

impl<'a> From<nom::Err<nom::error::Error<&'a [u8]>>> for OdsError {
    fn from(err: nom::Err<nom::error::Error<&'a [u8]>>) -> Self {
        match err {
            nom::Err::Incomplete(_) => unreachable!(),
            nom::Err::Error(err) => OdsError::Nom(nom::error::Error::new(
                String::from_utf8_lossy(err.input).to_string(),
                err.code,
            )),
            nom::Err::Failure(err) => OdsError::Nom(nom::error::Error::new(
                String::from_utf8_lossy(err.input).to_string(),
                err.code,
            )),
        }
    }
}

impl From<CellRefError> for OdsError {
    fn from(err: CellRefError) -> OdsError {
        OdsError::CellRef(format!("{:?}", err))
    }
}

impl<'s> From<ParseOFError<'s>> for OdsError {
    fn from(err: ParseOFError<'s>) -> OdsError {
        OdsError::CellRef(format!("{:?}", err))
    }
}
