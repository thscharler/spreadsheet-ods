pub use read::read_ods;
pub use write::write_ods;

pub use crate::error::OdsError;

mod read;
mod write;
mod xmlwriter;
mod tmp2zip;
