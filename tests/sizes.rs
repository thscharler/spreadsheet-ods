use std::mem::size_of;

use chrono::{Duration, NaiveDateTime};

use spreadsheet_ods::text::TextTag;
use spreadsheet_ods::{Sheet, Value, WorkBook};

#[test]
pub fn sizes() {
    println!("WorkBook {}", size_of::<WorkBook>());
    println!("Sheet {}", size_of::<Sheet>());

    println!("(ucell,ucell) {}", size_of::<(u32, u32)>());

    println!("Value {}", size_of::<Value>());

    println!("bool {}", size_of::<bool>());
    println!("f64 {}", size_of::<f64>());
    println!("f64, [u8; 3] {}", size_of::<(f64, [u8; 3])>());
    println!("f64, String {}", size_of::<(f64, String)>());
    println!("[u8; 3] {}", size_of::<[u8; 3]>());
    println!("String {}", size_of::<String>());
    println!("Vec<TextTag> {}", size_of::<Vec<TextTag>>());
    println!("NaiveDateTime {}", size_of::<NaiveDateTime>());
    println!("Duration {}", size_of::<Duration>());
}
