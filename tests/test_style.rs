use spreadsheet_ods::{cm, pt};
use spreadsheet_ods::style::{AttrFoBreak, AttrFoKeepWithNext, AttrFontDecl, AttrMap, AttrStyleWritingMode, AttrTableCol, AttrTableRow, FontFaceDecl, FontPitch, Length, PageBreak, TextKeep, WritingMode};
use spreadsheet_ods::Style;

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
