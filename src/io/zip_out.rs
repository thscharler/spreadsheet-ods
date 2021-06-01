use std::fs::File;
use std::io::{Error as IOError, Write};
use std::path::Path;

use zip::result::ZipError;
use zip::write::FileOptions;
use zip::ZipWriter;

/// Reduced Interface for ZipWriter.
#[allow(dead_code)]
pub struct ZipOut {
    zip: ZipWriter<File>,
}

pub struct ZipWrite<'a> {
    write: &'a mut ZipWriter<File>,
}

#[allow(dead_code)]
impl ZipOut {
    pub fn new(zip_file: &Path) -> Result<Self, std::io::Error> {
        let f = File::create(zip_file)?;
        Ok(ZipOut {
            zip: ZipWriter::new(f),
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
    ) -> Result<ZipWrite, ZipError> {
        self.zip.start_file(name, options)?;
        Ok(ZipWrite {
            write: &mut self.zip,
        })
    }

    pub fn zip(mut self) -> Result<File, ZipError> {
        self.zip.finish()
    }
}

impl<'a> Write for ZipWrite<'a> {
    fn write(&mut self, buf: &[u8]) -> Result<usize, IOError> {
        self.write.write(buf)
    }

    fn flush(&mut self) -> Result<(), IOError> {
        self.write.flush()
    }
}
