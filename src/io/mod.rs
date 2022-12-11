pub(crate) mod filebuf;
pub(crate) mod parse;
pub(crate) mod read;
pub(crate) mod write;

mod tmp2zip;
mod xmlwriter;
mod zip_out;

const DUMP_XML: bool = true;
const DUMP_UNUSED: bool = false;
