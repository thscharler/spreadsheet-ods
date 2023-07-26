use std::fs::File;
use std::io::{Cursor, Error as IOError, Seek, Write};
use std::path::Path;

use zip::result::ZipError;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

/// Reduced Interface for ZipWriter.
#[allow(dead_code)]
pub(crate) struct ZipOut<W: Write + Seek> {
    zip: ZipWriter<W>,
    compression: CompressionMethod,
}

pub(crate) struct ZipWrite<'a, W: Write + Seek> {
    write: &'a mut ZipWriter<W>,
}

#[allow(dead_code)]
impl<W: Write + Seek> ZipOut<W> {
    pub(crate) fn new_file(zip_file: &Path) -> Result<ZipOut<File>, std::io::Error> {
        let f = File::create(zip_file)?;
        Ok(ZipOut {
            zip: ZipWriter::new(f),
            compression: CompressionMethod::Deflated,
        })
    }

    pub(crate) fn new_buf(buf: Vec<u8>) -> Result<ZipOut<Cursor<Vec<u8>>>, std::io::Error> {
        Ok(ZipOut {
            zip: ZipWriter::new(Cursor::new(buf)),
            compression: CompressionMethod::Deflated,
        })
    }

    pub(crate) fn new_to(write: W) -> Result<ZipOut<W>, std::io::Error> {
        Ok(ZipOut {
            zip: ZipWriter::new(write),
            compression: CompressionMethod::Deflated,
        })
    }

    pub(crate) fn new_buf_uncompressed(
        buf: Vec<u8>,
    ) -> Result<ZipOut<Cursor<Vec<u8>>>, std::io::Error> {
        Ok(ZipOut {
            zip: ZipWriter::new(Cursor::new(buf)),
            compression: CompressionMethod::Stored,
        })
    }

    pub(crate) fn add_directory<S: Into<String>>(
        &mut self,
        path: S,
        options: FileOptions,
    ) -> Result<(), ZipError> {
        self.zip.add_directory(path, options)
    }

    pub(crate) fn start_file<S: Into<String>>(
        &mut self,
        name: S,
        options: FileOptions,
    ) -> Result<ZipWrite<'_, W>, ZipError> {
        let options = options.compression_method(self.compression);
        self.zip.start_file(name, options)?;
        Ok(ZipWrite {
            write: &mut self.zip,
        })
    }

    pub(crate) fn zip(mut self) -> Result<W, ZipError> {
        self.zip.finish()
    }
}

impl<W: Write + Seek> Write for ZipWrite<'_, W> {
    fn write(&mut self, buf: &[u8]) -> Result<usize, IOError> {
        self.write.write(buf)
    }

    fn flush(&mut self) -> Result<(), IOError> {
        self.write.flush()
    }
}
