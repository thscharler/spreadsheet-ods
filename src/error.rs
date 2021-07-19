use std::fmt::{Display, Formatter};

#[derive(Debug)]
pub enum OdsError {
    Ods(String),
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    Utf8(std::str::Utf8Error),
    Parse(String),
    ParseInt(std::num::ParseIntError),
    ParseBool(std::str::ParseBoolError),
    ParseFloat(std::num::ParseFloatError),
    Chrono(chrono::format::ParseError),
    Duration(time::OutOfRangeError),
    SystemTime(std::time::SystemTimeError),
}

impl Display for OdsError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            OdsError::Ods(e) => write!(f, "Ods {}", e)?,
            OdsError::Io(e) => write!(f, "IO {}", e)?,
            OdsError::Zip(e) => write!(f, "Zip {}", e)?,
            OdsError::Xml(e) => write!(f, "Xml {}", e)?,
            OdsError::Parse(e) => write!(f, "Parse {}", e)?,
            OdsError::ParseInt(e) => write!(f, "ParseInt {}", e)?,
            OdsError::ParseBool(e) => write!(f, "ParseBool {}", e)?,
            OdsError::ParseFloat(e) => write!(f, "ParseFloat {}", e)?,
            OdsError::Chrono(e) => write!(f, "Chrono {}", e)?,
            OdsError::Duration(e) => write!(f, "Duration {}", e)?,
            OdsError::SystemTime(e) => write!(f, "SystemTime {}", e)?,
            OdsError::Utf8(e) => write!(f, "UTF8 {}", e)?,
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
            OdsError::Parse(_) => None,
            OdsError::ParseInt(e) => Some(e),
            OdsError::ParseBool(e) => Some(e),
            OdsError::ParseFloat(e) => Some(e),
            OdsError::Chrono(e) => Some(e),
            OdsError::Duration(e) => Some(e),
            OdsError::SystemTime(e) => Some(e),
            OdsError::Utf8(e) => Some(e),
        }
    }
}

impl From<time::OutOfRangeError> for OdsError {
    fn from(err: time::OutOfRangeError) -> OdsError {
        OdsError::Duration(err)
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
