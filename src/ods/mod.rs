
pub use read::read_ods;
pub use write::write_ods_clean;
pub use write::write_ods;
pub use write2::write2_ods_clean;
pub use write2::write2_ods;
pub use error::OdsError;

mod xml_util;
mod temp_zip;

mod read;
mod write;
mod write2;

pub mod error;