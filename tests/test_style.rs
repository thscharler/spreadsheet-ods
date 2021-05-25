use color::Rgb;

use spreadsheet_ods::style::stylemap::StyleMap;
use spreadsheet_ods::style::{CellStyle, StyleOrigin, StyleUse, TableStyle};
use spreadsheet_ods::{write_ods, CellRef, OdsError, Sheet, WorkBook};

#[test]
fn testtablestyle() {
    let mut s = TableStyle::new("fine");
    s.set_background_color(Rgb::new(0, 0, 0));
}

#[test]
fn teststyles() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut ce12 = CellStyle::new("ce12", &"num2".into());
    ce12.set_origin(StyleOrigin::Styles);
    ce12.set_styleuse(StyleUse::Named);
    ce12.set_display_name("CC12");
    ce12.set_color(Rgb::new(192, 128, 0));
    wb.add_cellstyle(ce12);

    let mut ce11 = CellStyle::new("ce11", &"num2".into());
    ce11.set_origin(StyleOrigin::Styles);
    ce11.set_styleuse(StyleUse::Named);
    ce11.set_display_name("CC11");
    ce11.set_color(Rgb::new(0, 192, 128));
    wb.add_cellstyle(ce11);

    let mut ce13 = CellStyle::new("ce13", &"num4".into());
    ce13.push_stylemap(StyleMap::new(
        "cell-content()=\"BB\"",
        "ce12",
        CellRef::remote("s0", 4, 3),
    ));
    ce13.push_stylemap(StyleMap::new(
        "cell-content()=\"CC\"",
        "ce11",
        CellRef::remote("s0", 4, 3),
    ));
    let ce13 = wb.add_cellstyle(ce13);

    let mut sh = Sheet::new_with_name("s0");
    sh.set_styled_value(4, 3, "AA", &ce13);
    sh.set_styled_value(5, 3, "BB", &ce13);
    sh.set_styled_value(6, 3, "CC", &ce13);

    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/styles.ods")?;

    Ok(())
}
