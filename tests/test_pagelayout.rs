use color::Rgb;

use spreadsheet_ods::{cm, Length, OdsError, read_ods, WorkBook, write_ods};
use spreadsheet_ods::style::{AttrFoBackgroundColor, AttrFoMargin, AttrFoMinHeight, PageLayout};

#[test]
fn pagelayout() -> Result<(), OdsError> {
    let ods = read_ods("test_out/experiment.ods")?;
    //println!("{:?}", ods.pagelayout("Mpm1").unwrap().header().left());
    write_ods(&ods, "test_out/rexp.ods")?;

    Ok(())
}

#[test]
fn crpagelayout() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut pl = PageLayout::new_default();

    pl.set_background_color(Rgb::new(12, 129, 252));

    pl.header_attr_mut().set_min_height(cm!(0.75));
    pl.header_attr_mut().set_margin_left(cm!(0.15));
    pl.header_attr_mut().set_margin_right(cm!(0.15));
    pl.header_attr_mut().set_margin_bottom(cm!(0.15));

    pl.header_mut().center_mut().push_text("middle ground");
    pl.header_mut().left_mut().push_text("left wing");
    pl.header_mut().right_mut().push_text("right wing");

    wb.add_pagelayout(pl);

    write_ods(&wb, "test_out/hf0.ods")?;

    Ok(())
}
