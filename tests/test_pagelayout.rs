use color::Rgb;

use spreadsheet_ods::style::units::Length;
use spreadsheet_ods::style::{MasterPage, PageStyle, TableStyle};
use spreadsheet_ods::{cm, read_ods, write_ods, OdsError, Sheet, WorkBook};

#[test]
fn test_pagelayout() -> Result<(), OdsError> {
    let path = std::path::Path::new("test_out/format.ods");
    let mut ods;

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
        write_ods(&mut ods, path)?;
    } else {
        std::fs::create_dir_all(path.parent().unwrap())?;
        std::fs::File::create(path)?;
        write_ods(&mut ods, path)?;
    }

    Ok(())
}

#[test]
fn test_crpagelayout() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut ps = PageStyle::new("ps1");
    ps.set_background_color(Rgb::new(12, 129, 252));
    ps.headerstyle_mut().set_min_height(cm!(0.75));
    ps.headerstyle_mut().set_margin_left(cm!(0.15));
    ps.headerstyle_mut().set_margin_right(cm!(0.15));
    ps.headerstyle_mut().set_margin_bottom(cm!(0.15));
    let ps = wb.add_pagestyle(ps);

    let mut mp = MasterPage::new("mp1");
    mp.set_pagestyle(&ps);
    mp.header_mut().center_mut().add_text("middle ground");
    mp.header_mut().left_mut().add_text("left wing");
    mp.header_mut().right_mut().add_text("right wing");
    let mp = wb.add_masterpage(mp);

    let mut ts = TableStyle::new("ts1");
    ts.set_master_page_name(&mp);
    let _ts = wb.add_tablestyle(ts);

    wb.push_sheet(Sheet::new("1"));

    write_ods(&mut wb, "test_out/hf0.ods")?;

    Ok(())
}
