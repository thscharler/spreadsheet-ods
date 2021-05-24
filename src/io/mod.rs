pub use read::read_ods;
pub use write::write_ods;

pub use crate::error::OdsError;

mod read;
mod tmp2zip;
mod write;
mod xmlwriter;
mod zip_out;
