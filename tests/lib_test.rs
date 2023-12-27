#![allow(dead_code, unreachable_pub)]

use spreadsheet_ods::{OdsError, WorkBook};
use std::fmt::{Display, Formatter};
use std::fs;
use std::hint::black_box;
use std::path::Path;
use std::time::Instant;

pub fn test_write_ods<P: AsRef<Path>>(book: &mut WorkBook, ods_path: P) -> Result<(), OdsError> {
    fs::create_dir_all("test_out")?;
    spreadsheet_ods::write_ods(book, ods_path)
}

#[derive(Clone, Debug)]
pub struct Timing {
    pub name: String,
    pub skip: usize,
    pub runs: usize,
    pub divider: u64,

    /// samples in ns. already divided by divider.
    pub samples: Vec<f64>,
}

impl Timing {
    pub fn name<S: Into<String>>(mut self, name: S) -> Self {
        self.name = name.into();
        self
    }

    pub fn skip(mut self, skip: usize) -> Self {
        self.skip = skip;
        self
    }

    pub fn runs(mut self, runs: usize) -> Self {
        self.runs = runs;
        self
    }

    pub fn divider(mut self, divider: u64) -> Self {
        self.divider = divider;
        self
    }

    pub fn run<E, R>(&mut self, mut fun: impl FnMut() -> Result<R, E>) -> Result<R, E> {
        assert!(self.runs > 0);
        assert!(self.divider > 0);

        let mut bench = move || {
            let now = Instant::now();
            let result = fun();
            (now.elapsed(), result)
        };

        let mut samples_vec = Vec::with_capacity(self.runs);
        let mut n = 0;
        let result = loop {
            let (elapsed, result) = black_box(bench());
            samples_vec.push(elapsed);
            n += 1;
            if n >= self.runs + self.skip {
                break result;
            }
        };

        self.samples.extend(
            samples_vec
                .iter()
                .skip(self.skip)
                .map(|v| v.as_nanos() as f64 / self.divider as f64),
        );

        result
    }

    pub fn mean(&self) -> f64 {
        self.samples.iter().sum::<f64>() / self.samples.len() as f64
    }

    pub fn lin_dev(&self) -> f64 {
        let mean = self.mean();
        let lin_sum = self.samples.iter().map(|v| (*v - mean).abs()).sum::<f64>();
        lin_sum / self.samples.len() as f64
    }

    pub fn std_dev(&self) -> f64 {
        let mean = self.mean();
        let std_sum = self
            .samples
            .iter()
            .map(|v| (*v - mean) * (*v - mean))
            .sum::<f64>();
        (std_sum / self.samples.len() as f64).sqrt()
    }
}

impl Default for Timing {
    fn default() -> Self {
        Self {
            name: "".to_string(),
            skip: 0,
            runs: 1,
            divider: 1,
            samples: vec![],
        }
    }
}

impl Display for Timing {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        writeln!(f,)?;
        writeln!(f, "{}", self.name)?;
        writeln!(f,)?;
        writeln!(f, "| mean | lin_dev | std_dev |")?;
        writeln!(f, "|:---|:---|:---|")?;
        writeln!(
            f,
            "| {:.2} | {:.2} | {:.2} |",
            self.mean(),
            self.lin_dev(),
            self.std_dev()
        )?;
        writeln!(f,)?;

        if f.alternate() {
            writeln!(f,)?;
            writeln!(f, "{}", self.name)?;
            writeln!(f,)?;
            for i in 0..self.samples.len() {
                write!(f, "| {} ", i)?;
            }
            writeln!(f, "|")?;
            for _ in 0..self.samples.len() {
                write!(f, "|:---")?;
            }
            writeln!(f, "|")?;
            for e in &self.samples {
                write!(f, "| {:.2} ", e)?;
            }
            writeln!(f, "|")?;
        }

        Ok(())
    }
}
