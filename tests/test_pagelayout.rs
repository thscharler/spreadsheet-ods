use color::Rgb;

use spreadsheet_ods::style::PageLayout;
use spreadsheet_ods::{cm, read_ods, write_ods, Length, OdsError, WorkBook};

#[test]
fn pagelayout() -> Result<(), OdsError> {
    let path = std::path::Path::new("test_out/format.ods");
    let ods;

    if path.exists() {
        ods = read_ods(path)?;
    } else {
        std::fs::create_dir_all(path.parent().unwrap())?;
        std::fs::File::create(path)?;
        ods = read_ods(path)?;
    }
    //println!("{:?}", ods.pagelayout("Mpm1").unwrap().header().left());
    let path = std::path::Path::new("test_out/rexp.ods");

    if path.exists() {
        write_ods(&ods, path)?;
    } else {
        std::fs::create_dir_all(path.parent().unwrap())?;
        std::fs::File::create(path)?;
        write_ods(&ods, path)?;
    }

    Ok(())
}

#[test]
fn crpagelayout() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut pl = PageLayout::new_default();

    pl.set_background_color(Rgb::new(12, 129, 252));

    pl.header_style_mut().set_min_height(cm!(0.75));
    pl.header_style_mut().set_margin_left(cm!(0.15));
    pl.header_style_mut().set_margin_right(cm!(0.15));
    pl.header_style_mut().set_margin_bottom(cm!(0.15));

    pl.header_mut().center_mut().push_text("middle ground");
    pl.header_mut().left_mut().push_text("left wing");
    pl.header_mut().right_mut().push_text("right wing");

    wb.add_pagelayout(pl);

    write_ods(&wb, "test_out/hf0.ods")?;

    Ok(())
}
