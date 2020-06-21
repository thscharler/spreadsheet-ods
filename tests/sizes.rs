use std::mem::size_of;
use std::time::Duration;

use chrono::NaiveDateTime;

use spreadsheet_ods::{CellRange, CellRef, SCell, Sheet, Style, ucell, Value, ValueFormat, ValueType, WorkBook};
use spreadsheet_ods::format::FormatPart;
use spreadsheet_ods::style::{AttrMapType, HeaderFooter, PageLayout};
use spreadsheet_ods::text::{TextElem, TextTag, TextVec};

#[test]
pub fn sizes() {
    println!("WorkBook {}", size_of::<WorkBook>());
    println!("Sheet {}", size_of::<Sheet>());

    println!("SCell {}", size_of::<SCell>());
    println!("Value {}", size_of::<Value>());
    println!("Option<String> {}", size_of::<Option<String>>());
    println!("(ucell,ucell) {}", size_of::<(ucell, ucell)>());


    println!("bool {}", size_of::<bool>());
    println!("f64 {}", size_of::<f64>());
    println!("(String, f64) {}", size_of::<(String, f64)>());
    println!("String {}", size_of::<String>());
    println!("(String) {}", size_of::<String>());
    println!("TextVec {}", size_of::<TextVec>());
    println!("TextTag {}", size_of::<TextTag>());
    println!("Option<Vec<TextElem>> {}", size_of::<Option<Vec<TextElem>>>());
    println!("Vec<TextElem> {}", size_of::<Vec<TextElem>>());
    println!("Option<AttrMapType> {}", size_of::<Option<AttrMapType>>());
    println!("AttrMapType {}", size_of::<AttrMapType>());
    println!("NaiveDateTime {}", size_of::<NaiveDateTime>());
    println!("Duration {}", size_of::<Duration>());
    println!("Box<(String, f64)> {}", size_of::<Box<(String, f64)>>());


    println!("Value {}", size_of::<Value>());
    println!("ValueType {}", size_of::<ValueType>());
    println!("ValueFormat {}", size_of::<ValueFormat>());
    println!("FormatPart {}", size_of::<FormatPart>());
    println!("PageLayout {}", size_of::<PageLayout>());
    println!("HeaderFooter {}", size_of::<HeaderFooter>());
    println!("Style {}", size_of::<Style>());
    println!("CellRange {}", size_of::<CellRange>());
    println!("CellRef {}", size_of::<CellRef>());
}