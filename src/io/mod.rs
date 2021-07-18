pub use filebuf::*;
pub use read::default_settings;
pub use read::read_ods;
pub use read::read_ods_buf;
pub use write::write_ods;
pub use write::write_ods_buf;
pub use write::write_ods_buf_uncompressed;

pub use crate::error::OdsError;

mod filebuf;
mod read;
mod tmp2zip;
mod write;
mod xmlwriter;
mod zip_out;
