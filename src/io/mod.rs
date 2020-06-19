pub use read::read_ods;
pub use write::write_ods;
pub use write::write_ods_flags;

pub use crate::error::OdsError;

mod xmlwriter;
mod tmp2zip;

mod read;
mod write;
