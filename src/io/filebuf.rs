use std::fmt::{Debug, Formatter};

/// Directory or file.
#[derive(Clone)]
pub enum FileBufEntry {
    Dir(String),
    File(String, Vec<u8>),
}

impl Debug for FileBufEntry {
    fn fmt(&self, f: &mut Formatter<'_>) -> core::fmt::Result {
        match self {
            FileBufEntry::Dir(dir) => {
                write!(f, "{}", dir)?;
            }
            FileBufEntry::File(file, _) => {
                write!(f, "{}", file)?;
            }
        }
        Ok(())
    }
}

/// Acts as a buffer for files and directories.
#[derive(Clone, Debug)]
pub struct FileBuf {
    buf: Vec<FileBufEntry>,
}

impl Default for FileBuf {
    fn default() -> Self {
        FileBuf::new()
    }
}

impl FileBuf {
    pub fn new() -> Self {
        Self { buf: Vec::new() }
    }

    pub fn iter(&self) -> core::slice::Iter<FileBufEntry> {
        self.buf.iter()
    }

    pub fn contains<S: AsRef<str>>(&self, name: S) -> bool {
        for it in &self.buf {
            let found = match it {
                FileBufEntry::Dir(n) => n == name.as_ref(),
                FileBufEntry::File(n, _) => n == name.as_ref(),
            };

            if found {
                return true;
            }
        }

        false
    }

    pub fn push_dir<S: Into<String>>(&mut self, dir: S) {
        self.buf.push(FileBufEntry::Dir(dir.into()));
    }

    pub fn push_file<S: Into<String>>(&mut self, file: S, data: Vec<u8>) {
        self.buf.push(FileBufEntry::File(file.into(), data));
    }
}
