use std::mem::size_of;

use spreadsheet_ods::{Sheet, WorkBook};
use spreadsheet_ods::style::{FontFaceDecl, HeaderFooterAttr, PageLayout, Style, TextAttr};

#[test]
fn teststyles() {
    println!("WorkBook {}", size_of::<WorkBook>());
    println!("Sheet {}", size_of::<Sheet>());
    println!("Style {}", size_of::<Style>());
    println!("PageLayout {}", size_of::<PageLayout>());
    println!("TextAttr {}", size_of::<TextAttr>());
    println!("HeaderFooterAttr {}", size_of::<HeaderFooterAttr>());
    println!("FontFaceDecl {}", size_of::<FontFaceDecl>());
}