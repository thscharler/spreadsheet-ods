use std::mem::size_of;

use chrono::{Duration, NaiveDateTime};

use spreadsheet_ods::text::TextTag;
use spreadsheet_ods::{Sheet, Value, WorkBook};

#[test]
pub fn sizes() {
    println!("WorkBook {}", size_of::<WorkBook>());
    println!("Sheet {}", size_of::<Sheet>());

    println!("Value {}", size_of::<Value>());
    println!("Option<String> {}", size_of::<Option<String>>());
    println!("(ucell,ucell) {}", size_of::<(u32, u32)>());

    println!("bool {}", size_of::<bool>());
    println!("f64 {}", size_of::<f64>());
    println!("String, f64 {}", size_of::<(String, f64)>());
    println!("String {}", size_of::<String>());
    println!("Vec<TextTag> {}", size_of::<Vec<TextTag>>());
    println!("NaiveDateTime {}", size_of::<NaiveDateTime>());
    println!("Duration {}", size_of::<Duration>());

    // println!("bool {}", size_of::<bool>());
    // println!("f64 {}", size_of::<f64>());
    // println!("(String, f64) {}", size_of::<(String, f64)>());
    // println!("String {}", size_of::<String>());
    // println!("(String) {}", size_of::<String>());
    // println!("TextTag {}", size_of::<TextTag>());
    // println!("Box<TextTag> {}", size_of::<Box<TextTag>>());
    // println!("XmlContent {}", size_of::<XmlContent>());
    // println!("Vec<XmlContent> {}", size_of::<Vec<XmlContent>>());
    // println!("Option<AttrMapType> {}", size_of::<Option<AttrMapType>>());
    // println!("AttrMapType {}", size_of::<AttrMapType>());
    // println!("NaiveDateTime {}", size_of::<NaiveDateTime>());
    // println!("Duration {}", size_of::<Duration>());
    // println!("Box<(String, f64)> {}", size_of::<Box<(String, f64)>>());
    //
    // println!("Value {}", size_of::<Value>());
    // println!("ValueType {}", size_of::<ValueType>());
    // println!("ValueFormat {}", size_of::<ValueFormat>());
    // println!("FormatPart {}", size_of::<FormatPart>());
    // println!("PageLayout {}", size_of::<PageLayout>());
    // println!("HeaderFooter {}", size_of::<HeaderFooter>());
    // println!("Style {}", size_of::<Style>());
    // println!("TableAttr {}", size_of::<TableAttr>());
    // println!("ParagraphAttr {}", size_of::<ParagraphAttr>());
    // println!("Vec<StyleMap> {}", size_of::<Vec<StyleMap>>());
    // println!("CellRange {}", size_of::<CellRange>());
    // println!("CellRef {}", size_of::<CellRef>());
}
