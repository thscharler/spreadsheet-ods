
pub use read::read_ods;
pub use write::write_ods_clean;
pub use write::write_ods;
pub use error::OdsError;

mod xmlwriter;
mod tmp2zip;

mod read;
mod write;

pub mod error;