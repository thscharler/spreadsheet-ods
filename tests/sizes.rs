use std::alloc::{alloc_zeroed, Layout};
use std::collections::HashMap;
use std::mem;
use std::mem::size_of;
use std::sync::Arc;

use chrono::{Duration, NaiveDateTime};
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

#[test]
pub fn smol() {
    println!("SmolStr {}", size_of::<SmolStr>());
    println!("Option<SmolStr> {}", size_of::<Option<SmolStr>>());
    println!(
        "mem value len 10 {}",
        mem::size_of_val(&SmolStr::from("1234567890"))
    );
    println!(
        "mem value len 30 {}",
        mem::size_of_val(&SmolStr::from("123456789012345678901234567890"))
    );
    println!(
        "loupe value len 10 {}",
        size_of_smolstr(&SmolStr::from("1234567890"))
    );
    println!(
        "loupe value len 30 {}",
        size_of_smolstr(&SmolStr::from("123456789012345678901234567890"))
    );
    println!(
        "loupe arc len 30 {}",
        loupe::size_of_val(&Arc::new("123456789012345678901234567890"))
    );
    println!(
        "mem arc len 30 {}",
        mem::size_of_val(&Arc::new("123456789012345678901234567890"))
    );

    let layout = Layout::new::<Arc<str>>().pad_to_align();
    println!("layout {:?}", layout);

    let layout = Layout::new::<SmolStr>().pad_to_align();
    println!("layout {:?}", layout);

    let layout = Layout::new::<[u8; 30]>().pad_to_align();
    println!("layout {:?}", layout);
}

pub(crate) fn size_of_smolstr(str: &SmolStr) -> usize {
    if str.is_heap_allocated() {
        mem::size_of_val(str) + str.len()
    } else {
        mem::size_of_val(str)
    }
}
