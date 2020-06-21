use std::mem::size_of;

use spreadsheet_ods::{CellRange, CellRef, SCell, Sheet, Style, Value, ValueFormat, ValueType, WorkBook};
use spreadsheet_ods::format::FormatPart;
use spreadsheet_ods::style::{HeaderFooter, PageLayout};

#[test]
pub fn sizes() {
    println!("WorkBook {}", size_of::<WorkBook>());
    println!("Sheet {}", size_of::<Sheet>());
    println!("SCell {}", size_of::<SCell>());
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