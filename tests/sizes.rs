use std::collections::HashMap;
use std::mem::size_of;

use chrono::{Duration, NaiveDate, NaiveDateTime};

use spreadsheet_ods::metadata::Metadata;
use spreadsheet_ods::style::TableStyle;
use spreadsheet_ods::text::TextTag;
use spreadsheet_ods::{Sheet, Value, WorkBook};

#[test]
pub fn sizes() {
    println!("WorkBook {}", size_of::<WorkBook>());
    println!("Sheet {}", size_of::<Sheet>());
    println!("Metadata {}", size_of::<Metadata>());

    println!("(ucell,ucell) {}", size_of::<(u32, u32)>());

    println!("Value {}", size_of::<Value>());

    println!("bool {}", size_of::<bool>());
    println!("f64 {}", size_of::<f64>());
    println!("f64, [u8; 3] {}", size_of::<(f64, [u8; 3])>());
    println!("f64, String {}", size_of::<(f64, String)>());
    println!("[u8; 3] {}", size_of::<[u8; 3]>());
    println!("String {}", size_of::<String>());
    println!("Vec<TextTag> {}", size_of::<Vec<TextTag>>());
    println!(
        "HashMap<String, TableStyle> {}",
        size_of::<HashMap<String, TableStyle>>()
    );
    println!("NaiveDateTime {}", size_of::<NaiveDateTime>());
    println!("Duration {}", size_of::<Duration>());
}

#[test]
fn test_size2() {
    #[repr(u8)]
    pub enum Value {
        Empty,
        Boolean(bool),
        Number(f64),
        Percentage(f64),
        Currency(f64),
        Text(Box<str>),
        TextXml(Box<[TextTag]>),
        DateTime(f64),
        TimeDuration(f64),
    }

    println!("Value {}", size_of::<spreadsheet_ods::Value>());
    println!("Value {}", size_of::<Value>());
    println!("f64 {}", size_of::<f64>());
    println!("Option<f64> {}", size_of::<Option<f64>>());
    println!("Box<str> {}", size_of::<Box<str>>());
    println!("Box<[TextTag]> {}", size_of::<Box<[TextTag]>>());
}

#[test]
fn test_f64_date() {
    fn to_f64(dt: NaiveDateTime) -> f64 {
        0.0
    }

    fn from_f64(fdt: f64) -> Option<NaiveDateTime> {
        let secs_f = fdt.floor();
        let secs = secs_f as i64;
        let nsecs = ((fdt - secs_f) * 1e9) as u32;
        NaiveDateTime::from_timestamp_opt(secs, nsecs)
    }
}
