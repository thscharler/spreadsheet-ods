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
use std::io::{BufWriter, Read, Write};
use std::path::{Path, PathBuf};

pub struct TempZip {
    zipped: PathBuf,
    temp_path: PathBuf,
    temp_file: Option<File>,
    entries: Vec<TempZipEntry>,
}

struct TempZipEntry {
    is_dir: bool,
    name: String,
    fopt: zip::write::FileOptions,
}

impl TempZip {
    /// Final ZIP is this file. The temporaries are written at the
    /// same location, in a new directory with the same basename.
    pub fn new(zip_file: &Path) -> Self {
        let mut path: PathBuf = zip_file.parent().unwrap().to_path_buf();
        path.push(zip_file.file_stem().unwrap());

        TempZip {
            zipped: zip_file.to_path_buf(),
            temp_path: path,
            temp_file: None,
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
    pub fn start_file(&mut self, name: &str, fopt: zip::write::FileOptions) -> Result<(), std::io::Error> {
        let file = self.temp_path.join(name);
        let path = file.parent().unwrap();
        create_dir_all(path)?;

        self.temp_file = Some(File::create(file)?);

        self.entries.push(TempZipEntry {
            is_dir: false,
            name: name.to_string(),
            fopt,
        });

        Ok(())
    }

    /// Packs all created files into a zip.
    pub fn zip(&mut self) -> Result<(), std::io::Error> {
        // should close the last file?
        self.temp_file = None;

        let zip_file = File::create(&self.zipped)?;
        let mut zip_writer = zip::ZipWriter::new(BufWriter::new(zip_file));

        for entry in &self.entries {
            if !entry.is_dir {
                zip_writer.start_file(&entry.name, entry.fopt)?;

                let file = self.temp_path.join(&entry.name);
                let mut rd = File::open(file)?;
                let mut buf = vec![];
                rd.read_to_end(&mut buf)?;

                zip_writer.write_all(buf.as_slice())?;
            } else {
                zip_writer.add_directory(&entry.name, entry.fopt)?;
            }
        }

        Ok(())
    }

    // Cleanup
    pub fn clean(&mut self) -> Result<(), std::io::Error> {
        remove_dir_all(&self.temp_path)?;
        Ok(())
    }
}

impl Write for TempZip {
    fn write(&mut self, buf: &[u8]) -> Result<usize, std::io::Error> {
        if let Some(file) = &mut self.temp_file {
            file.write(buf)
        } else {
            panic!("No file to write");
        }
    }

    fn flush(&mut self) -> Result<(), std::io::Error> {
        if let Some(file) = &mut self.temp_file {
            file.flush()
        } else {
            panic!("No file to write");
        }
    }
}
