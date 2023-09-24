pub(crate) mod format;
pub(crate) mod parse;
pub(crate) mod read;
pub(crate) mod write;

mod tmp2zip;
mod xmlwriter;
mod zip_out;

const DUMP_XML: bool = false;
const DUMP_UNUSED: bool = false;
