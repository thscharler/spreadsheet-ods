//use spreadsheet_ods::attrmap::{AttrMap, AttrFoBackgroundColor, AttrFoBorder, Border, AttrFoMargin, AttrFoPadding, AttrStyleDynamicSpacing, AttrStyleShadow, AttrSvgHeight, AttrFoMinHeight, Length};
use color::Rgb;

use spreadsheet_ods::{cm, mm, pt};
use spreadsheet_ods::style::{AttrFoBackgroundColor, AttrFoBorder, AttrFoMargin, AttrFoMinHeight, AttrFoPadding, AttrMap, AttrStyleDynamicSpacing, AttrStyleShadow, AttrSvgHeight, Border, Length, PageLayout};

#[test]
fn test_attr1() {
    let mut p0 = PageLayout::default();

    p0.set_background_color(Rgb::new(12, 33, 46));
    assert_eq!(p0.attr("fo:background-color"), Some(&"#0c212e".to_string()));

    p0.set_border(pt!(1), Border::Groove, Rgb::new(99, 0, 0));
    assert_eq!(p0.attr("fo:border"), Some(&"1pt groove #630000".to_string()));

    p0.set_border_line_width(pt!(1), pt!(2), pt!(3));
    assert_eq!(p0.attr("style:border-line-width"), Some(&"1pt 2pt 3pt".to_string()));

    p0.set_margin(Length::Pt(3.2));
    assert_eq!(p0.attr("fo:margin"), Some(&"3.2pt".to_string()));

    p0.set_margin(pt!(3.2));
    assert_eq!(p0.attr("fo:margin"), Some(&"3.2pt".to_string()));

    p0.set_padding(pt!(3.3));
    assert_eq!(p0.attr("fo:padding"), Some(&"3.3pt".to_string()));

    p0.set_dynamic_spacing(true);
    assert_eq!(p0.attr("style:dynamic-spacing"), Some(&"true".to_string()));

    p0.set_shadow(mm!(3), mm!(4), None, Rgb::new(16, 16, 16));
    assert_eq!(p0.attr("style:shadow"), Some(&"#101010 3mm 4mm".to_string()));

    p0.set_height(cm!(7));
    assert_eq!(p0.attr("svg:height"), Some(&"7cm".to_string()));

    p0.header_attr_mut().set_min_height(cm!(6));
    assert_eq!(p0.header_attr().attr("fo:min-height"), Some(&"6cm".to_string()));

    p0.header_attr_mut().set_dynamic_spacing(true);
    assert_eq!(p0.header_attr().attr("style:dynamic-spacing"), Some(&"true".to_string()));
}