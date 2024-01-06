use std::alloc::{alloc_zeroed, Layout};
use std::collections::HashMap;
use std::mem;
use std::mem::size_of;
use std::sync::Arc;

use chrono::{Duration, NaiveDateTime};
use get_size::GetSize;
use get_size_derive::GetSize;
use smol_str::SmolStr;

use spreadsheet_ods::metadata::Metadata;
use spreadsheet_ods::style::TableStyle;
use spreadsheet_ods::text::TextTag;
use spreadsheet_ods::{Sheet, Value, WorkBook};

// #[test]
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

struct S0 {
    v0: SmolStr,
    v1: SmolStr,
}

impl GetSize for S0 {
    fn get_heap_size(&self) -> usize {
        heap_size_of_smolstr(&self.v0) + heap_size_of_smolstr(&self.v1)
    }
}

struct S1 {
    v0: String,
    v1: String,
}

impl GetSize for S1 {
    fn get_heap_size(&self) -> usize {
        let mut total = 0;
        total += GetSize::get_heap_size(&self.v0);
        total += GetSize::get_heap_size(&self.v1);
        total
    }
}

#[test]
pub fn smol() {
    println!("SmolStr {}", size_of::<SmolStr>());
    println!("Option<SmolStr> {}", size_of::<Option<SmolStr>>());
    println!(
        "loupe SmolStr len 10 {}",
        size_of_smolstr(&SmolStr::from("1234567890"))
    );
    println!(
        "loupe SmolStr len 30 {}",
        size_of_smolstr(&SmolStr::from("123456789012345678901234567890"))
    );
    println!(
        "loupe Arc len 30 {}",
        Arc::new("123456789012345678901234567890").get_size()
    );
    println!(
        "loupe String len 30 {}",
        String::from("123456789012345678901234567890").get_size()
    );

    let v0 = S0 {
        v0: SmolStr::from("123456789012345678901234567890"),
        v1: SmolStr::from("123456789012345678901234567890"),
    };
    println!("v0 {}", v0.get_size() / 2);
    let v1 = S1 {
        v0: "123456789012345678901234567890".to_string(),
        v1: "123456789012345678901234567890".to_string(),
    };
    println!("v1 {}", v1.get_size() / 2);

    let v0 = S0 {
        v0: SmolStr::from("1234567890"),
        v1: SmolStr::from("1234567890"),
    };
    println!("v0 {}", v0.get_size() / 2);
    let v1 = S1 {
        v0: "1234567890".to_string(),
        v1: "1234567890".to_string(),
    };
    println!("v1 {}", v1.get_size() / 2);
}

pub(crate) fn size_of_smolstr(str: &SmolStr) -> usize {
    if str.is_heap_allocated() {
        mem::size_of_val(str) + str.len()
    } else {
        mem::size_of_val(str)
    }
}

pub(crate) fn heap_size_of_smolstr(str: &SmolStr) -> usize {
    if str.is_heap_allocated() {
        str.len()
    } else {
        0
    }
}
