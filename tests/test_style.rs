use color::Rgb;

use spreadsheet_ods::style::{StyleMap, StyleOrigin, StyleUse, TableCellStyle, TableStyle};
use spreadsheet_ods::{write_ods, CellRef, OdsError, Sheet, WorkBook};

#[test]
fn testtablestyle() {
    let mut s = TableStyle::new("fine");
    s.set_background_color(Rgb::new(0, 0, 0));
}

#[test]
fn teststyles() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut st = TableCellStyle::new("ce12", "num2");
    st.set_origin(StyleOrigin::Styles);
    st.set_styleuse(StyleUse::Named);
    st.set_display_name("CC12");
    st.set_color(Rgb::new(192, 128, 0));
    wb.add_cell_style(st);

    let mut st = TableCellStyle::new("ce11", "num2");
    st.set_origin(StyleOrigin::Styles);
    st.set_styleuse(StyleUse::Named);
    st.set_display_name("CC11");
    st.set_color(Rgb::new(0, 192, 128));
    wb.add_cell_style(st);

    let mut st = TableCellStyle::new("ce13", "num4");
    st.push_stylemap(StyleMap::new(
        "cell-content()=\"BB\"",
        "ce12",
        CellRef::remote("s0", 4, 3),
    ));
    st.push_stylemap(StyleMap::new(
        "cell-content()=\"CC\"",
        "ce11",
        CellRef::remote("s0", 4, 3),
    ));
    wb.add_cell_style(st);

    let mut sh = Sheet::new_with_name("s0");
    sh.set_styled_value(4, 3, "AA", "ce13");
    sh.set_styled_value(5, 3, "BB", "ce13");
    sh.set_styled_value(6, 3, "CC", "ce13");

    wb.push_sheet(sh);

    write_ods(&wb, "test_out/styles.ods")?;

    Ok(())
}
