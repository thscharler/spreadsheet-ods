/// Writing to the zip directly strangely broke my content.xml
/// Could'nt really find the source of this, so for now I use
/// this replacer. Simulates the same API but writes to a
/// temp-directory.
///
/// I zip all the stuff afterwards, and it's not to uncomfortable
/// to have the unzipped files around für debugging. z
///
/// zip_clean() cleans up afterwards.
///

use std::fs::{create_dir_all, File, remove_dir_all};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};
use std::collections::HashSet;

pub struct TempZip {
    zipped: PathBuf,
    temp_path: PathBuf,
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
    /// Final ZIP is this file. The temporaries are written at the
    /// same location, in a new directory with the same basename.
    pub fn new(zip_file: &Path) -> Self {
        let mut path: PathBuf = zip_file.parent().unwrap().to_path_buf();
        path.push(zip_file.file_stem().unwrap());

        if path.exists() {
            panic!("ZIP temp directory {:?} already exists!", path);
        }

        TempZip {
            zipped: zip_file.to_path_buf(),
            temp_path: path,
            entries: Vec::new(),
        }
    }

    /// Adds this directory.
    pub fn add_directory(&mut self, name: &str, fopt: zip::write::FileOptions) -> Result<(), std::io::Error> {
        let add = self.temp_path.join(name);

        create_dir_all(&add)?;

        self.entries.push(TempZipEntry {
            is_dir: true,
            name: name.to_string(),
            fopt,
        });

        Ok(())
    }

    /// Starts a new file inside the zip. After calling this function
    /// the Write trait for TempZip starts working.
    pub fn start_file<'a>(&'a mut self, name: &'_ str, fopt: zip::write::FileOptions) -> Result<TempWrite<'a>, std::io::Error> {
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

    // Cleanup
    pub fn clean(&mut self) -> Result<(), std::io::Error> {
        remove_dir_all(&self.temp_path)?;
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
