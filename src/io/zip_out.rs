use std::fs::File;
use std::io::{Cursor, Error as IOError, Seek, Write};
use std::path::Path;

use zip::result::ZipError;
use zip::write::FileOptions;
use zip::ZipWriter;

/// Reduced Interface for ZipWriter.
#[allow(dead_code)]
pub struct ZipOut<W: Write + Seek> {
    zip: ZipWriter<W>,
}

pub struct ZipWrite<'a, W: Write + Seek> {
    write: &'a mut ZipWriter<W>,
}

#[allow(dead_code)]
impl<W: Write + Seek> ZipOut<W> {
    pub fn new_file(zip_file: &Path) -> Result<ZipOut<File>, std::io::Error> {
        let f = File::create(zip_file)?;
        Ok(ZipOut {
            zip: ZipWriter::new(f),
        })
    }

    pub fn new_buf(buf: Vec<u8>) -> Result<ZipOut<Cursor<Vec<u8>>>, std::io::Error> {
        Ok(ZipOut {
            zip: ZipWriter::new(Cursor::new(buf)),
        })
    }

    pub fn add_directory<S: Into<String>>(
        &mut self,
        path: S,
        options: FileOptions,
    ) -> Result<(), ZipError> {
        self.zip.add_directory(path, options)
    }

    pub fn start_file<S: Into<String>>(
        &mut self,
        name: S,
        options: FileOptions,
    ) -> Result<ZipWrite<W>, ZipError> {
        self.zip.start_file(name, options)?;
        Ok(ZipWrite {
            write: &mut self.zip,
        })
    }

    pub fn zip(mut self) -> Result<W, ZipError> {
        self.zip.finish()
    }
}

impl<'a, W: Write + Seek> Write for ZipWrite<'a, W> {
    fn write(&mut self, buf: &[u8]) -> Result<usize, IOError> {
        self.write.write(buf)
    }

    fn flush(&mut self) -> Result<(), IOError> {
        self.write.flush()
    }
}
