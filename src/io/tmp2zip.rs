//! Writing to the zip directly strangely broke my content.xml
//! Could'nt really find the source of this, so for now I use
//! this replacer.

use mktemp::Temp;
use std::collections::HashSet;
use std::fs::{create_dir_all, File};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};

/// Data for the ZIP archive.
pub struct TempZip {
    zipped: PathBuf,
    temp_path: Temp,
    entries: Vec<TempZipEntry>,
}

pub struct TempWrite<'a> {
    _temp_zip: &'a TempZip,
    temp_file: BufWriter<File>,
}

struct TempZipEntry {
    name: String,
    is_dir: bool,
    fopt: zip::write::FileOptions,
}

impl TempZip {
    /// Final ZIP is this file. The temporary files are written to a
    /// new random subdirectory.
    pub fn new(zip_file: &Path) -> Result<Self, std::io::Error> {
        Ok(TempZip {
            zipped: zip_file.to_path_buf(),
            temp_path: mktemp::Temp::new_dir_in(zip_file.parent().unwrap())?,
            entries: Vec::new(),
        })
    }

    /// Adds this directory.
    pub fn add_directory(
        &mut self,
        name: &str,
        fopt: zip::write::FileOptions,
    ) -> Result<(), std::io::Error> {
        let add = self.temp_path.join(name);

        create_dir_all(&add)?;

        self.entries.push(TempZipEntry {
            is_dir: true,
            name: name.to_string(),
            fopt,
        });

        Ok(())
    }

    /// Starts a new file inside the zip. Returns a Write for the file.
    pub fn start_file<'a>(
        &'a mut self,
        name: &'_ str,
        fopt: zip::write::FileOptions,
    ) -> Result<TempWrite<'a>, std::io::Error> {
        let file = self.temp_path.join(name);
        let path = file.parent().unwrap();

        create_dir_all(path)?;

        self.entries.push(TempZipEntry {
            is_dir: false,
            name: name.to_string(),
            fopt,
        });

        Ok(TempWrite {
            _temp_zip: self,
            temp_file: BufWriter::new(File::create(file)?),
        })
    }

    /// Packs all created files into a zip.
    pub fn zip(&mut self) -> Result<(), std::io::Error> {
        let zip_file = File::create(&self.zipped)?;

        let mut zip_writer = zip::ZipWriter::new(BufWriter::new(zip_file));

        // deduplizieren.
        let mut names = HashSet::new();

        for entry in &self.entries {
            if names.contains(&entry.name) {
                // noop
            } else if !entry.is_dir {
                zip_writer.start_file(&entry.name, entry.fopt)?;

                let file = self.temp_path.join(&entry.name);
                let buf = std::fs::read(file)?;
                zip_writer.write_all(buf.as_slice())?;
            } else {
                zip_writer.add_directory(&entry.name, entry.fopt)?;
            }

            names.insert(entry.name.clone());
        }

        Ok(())
    }
}

impl<'a> Write for TempWrite<'a> {
    fn write(&mut self, buf: &[u8]) -> Result<usize, std::io::Error> {
        self.temp_file.write(buf)
    }

    fn flush(&mut self) -> Result<(), std::io::Error> {
        self.temp_file.flush()
    }
}
