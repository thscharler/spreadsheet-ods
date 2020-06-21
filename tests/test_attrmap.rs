use color::Rgb;

use spreadsheet_ods::{Angle, cm, deg, Length, mm, pt};
use spreadsheet_ods::style::{AttrFoBackgroundColor, AttrFoBorder, AttrFoMargin,
                             AttrFoMinHeight, AttrFoPadding, AttrMap, AttrStyleDynamicSpacing,
                             AttrStyleShadow, AttrSvgHeight, PageLayout};
use spreadsheet_ods::style::{AttrFoBreak, AttrFoKeepWithNext, AttrFontDecl, AttrParagraph,
                             AttrStyleWritingMode, AttrTableCell, AttrTableCol, AttrTableRow, AttrText,
                             Border, CellAlignVertical, FontFaceDecl, FontPitch, PageBreak, ParaAlignVertical,
                             RotationAlign, TextAlignSource, TextKeep, TextPosition,
                             TextRelief, TextTransform, TextWeight, WrapOption, WritingMode};
use spreadsheet_ods::Style;

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


#[test]
fn test_attr2() {
    let mut ff = FontFaceDecl::new();

    ff.set_font_family("Helvetica");
    assert_eq!(ff.attr("svg:font-family"), Some(&"Helvetica".to_string()));

    ff.set_font_family_generic("fool");
    assert_eq!(ff.attr("style:font-family-generic"), Some(&"fool".to_string()));

    ff.set_font_pitch(FontPitch::Fixed);
    assert_eq!(ff.attr("style:font-pitch"), Some(&"fixed".to_string()));
}

#[test]
fn test_attr3() {
    let mut st = Style::cell_style("c00", "f00");

    st.table_mut().set_break_before(PageBreak::Page);
    assert_eq!(st.table().attr("fo:break-before"), Some(&"page".to_string()));

    st.table_mut().set_break_after(PageBreak::Page);
    assert_eq!(st.table().attr("fo:break-after"), Some(&"page".to_string()));

    st.table_mut().set_keep_with_next(TextKeep::Auto);
    assert_eq!(st.table().attr("fo:keep-with-next"), Some(&"auto".to_string()));

    st.table_mut().set_writing_mode(WritingMode::TbLr);
    assert_eq!(st.table().attr("style:writing-mode"), Some(&"tb-lr".to_string()));

    st.col_mut().set_use_optimal_col_width(true);
    assert_eq!(st.col().attr("style:use-optimal-column-width"), Some(&"true".to_string()));

    st.col_mut().set_rel_col_width(33f32);
    assert_eq!(st.col().attr("style:rel-column-width"), Some(&"33*".to_string()));

    st.col_mut().set_col_width(cm!(17));
    assert_eq!(st.col().attr("style:column-width"), Some(&"17cm".to_string()));

    st.row_mut().set_use_optimal_row_height(true);
    assert_eq!(st.row().attr("style:use-optimal-row-height"), Some(&"true".to_string()));

    st.row_mut().set_min_row_height(pt!(77));
    assert_eq!(st.row().attr("style:min-row-height"), Some(&"77pt".to_string()));

    st.row_mut().set_row_height(pt!(77));
    assert_eq!(st.row().attr("style:row-height"), Some(&"77pt".to_string()));
}

#[test]
fn test_attr4() {
    let mut st = Style::cell_style("c00", "f00");

    st.cell_mut().set_diagonal_bl_tr(pt!(0.2), Border::Ridge, Rgb::new(0, 127, 0));
    assert_eq!(st.cell().attr("style:diagonal-bl-tr"), Some(&"0.2pt ridge #007f00".to_string()));

    st.cell_mut().set_diagonal_tl_br(pt!(0.2), Border::Ridge, Rgb::new(0, 127, 0));
    assert_eq!(st.cell().attr("style:diagonal-bl-tr"), Some(&"0.2pt ridge #007f00".to_string()));

    st.cell_mut().set_wrap_option(WrapOption::Wrap);
    assert_eq!(st.cell().attr("fo:wrap-option"), Some(&"wrap".to_string()));

    st.cell_mut().set_print_content(true);
    assert_eq!(st.cell().attr("style:print-content"), Some(&"true".to_string()));

    st.cell_mut().set_repeat_content(true);
    assert_eq!(st.cell().attr("style:repeat-content"), Some(&"true".to_string()));

    st.cell_mut().set_rotation_align(RotationAlign::Center);
    assert_eq!(st.cell().attr("style:rotation-align"), Some(&"center".to_string()));

    st.cell_mut().set_rotation_angle(deg!(42));
    assert_eq!(st.cell().attr("style:rotation-angle"), Some(&"42deg".to_string()));

    st.cell_mut().set_shrink_to_fit(true);
    assert_eq!(st.cell().attr("style:shrink-to-fit"), Some(&"true".to_string()));

    st.cell_mut().set_vertical_align(CellAlignVertical::Top);
    assert_eq!(st.cell().attr("style:vertical-align"), Some(&"top".to_string()));
}

#[test]
fn test_attr5() {
    let mut st = Style::cell_style("c00", "f00");

    st.paragraph_mut().set_vertical_align(ParaAlignVertical::Baseline);
    assert_eq!(st.paragraph().attr("style:vertical-align"), Some(&"baseline".to_string()));

    st.paragraph_mut().set_line_spacing(pt!(4));
    assert_eq!(st.paragraph().attr("style:line-spacing"), Some(&"4pt".to_string()));

    st.paragraph_mut().set_number_lines(true);
    assert_eq!(st.paragraph().attr("text:number-lines"), Some(&"true".to_string()));

    st.paragraph_mut().set_text_align_source(TextAlignSource::ValueType);
    assert_eq!(st.paragraph().attr("style:text-align-source"), Some(&"value-type".to_string()));

    st.paragraph_mut().set_text_indent(mm!(4.2));
    assert_eq!(st.paragraph().attr("fo:text-indent"), Some(&"4.2mm".to_string()));
}

#[test]
fn test_attr6() {
    let mut st = Style::cell_style("c00", "f00");

    st.text_mut().set_font_bold();
    assert_eq!(st.text().attr("fo:font-weight"), Some(&"bold".to_string()));

    st.text_mut().set_font_weight(TextWeight::W700);
    assert_eq!(st.text().attr("fo:font-weight"), Some(&"700".to_string()));

    st.text_mut().set_font_size(pt!(13));
    assert_eq!(st.text().attr("fo:font-size"), Some(&"13pt".to_string()));

    st.text_mut().set_color(Rgb::new(0, 0, 128));
    assert_eq!(st.text().attr("fo:color"), Some(&"#000080".to_string()));

    st.text_mut().set_font_italic();
    assert_eq!(st.text().attr("fo:font-style"), Some(&"italic".to_string()));

    st.text_mut().set_font_name("Boing");
    assert_eq!(st.text().attr("style:font-name"), Some(&"Boing".to_string()));

    st.text_mut().set_font_relief(TextRelief::Engraved);
    assert_eq!(st.text().attr("style:font-relief"), Some(&"engraved".to_string()));

    st.text_mut().set_letter_spacing(pt!(0.2));
    assert_eq!(st.text().attr("fo:letter-spacing"), Some(&"0.2pt".to_string()));

    st.text_mut().set_text_position(TextPosition::Sub);
    assert_eq!(st.text().attr("style:text-position"), Some(&"sub".to_string()));

    st.text_mut().set_text_transform(TextTransform::Lowercase);
    assert_eq!(st.text().attr("fo:text-transform"), Some(&"lowercase".to_string()));
}

