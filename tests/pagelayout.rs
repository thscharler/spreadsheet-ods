use color::Rgb;

use spreadsheet_ods::{Composit, CompositTag, CompositVec, OdsError, WorkBook};
use spreadsheet_ods::attrmap::{AttrFoBackground, AttrFoMargin, AttrFoMinHeight};
use spreadsheet_ods::io::{read_ods, write_ods};
use spreadsheet_ods::style::{HeaderFooter, PageLayout};

#[test]
fn pagelayout() -> Result<(), OdsError> {
    let ods = read_ods("test_out/experiment.ods")?;
    println!("{:?}", ods.pagelayout("Mpm1").unwrap().header().region_left());
    write_ods(&ods, "test_out/rexp.ods")?;

    Ok(())
}

#[test]
fn crpagelayout() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut pl = PageLayout::default();


    pl.set_background_color(Rgb::new(12, 129, 252));

    pl.header_attr().set_min_height("0.75cm");
    pl.header_attr().set_margin_left("0.15cm");
    pl.header_attr().set_margin_right("0.15cm");
    pl.header_attr().set_margin_bottom("0.15cm");

    //pl.set_prp("style:writing-mode", "lr-tb".to_string());

    let mut hf = HeaderFooter::new();
    let mut cv = CompositVec::new();
    cv.push(Composit::Start(CompositTag::new("text:p")));
    cv.push(Composit::Text("sioltard".to_string()));
    cv.push(Composit::End("text:p".to_string()));
    hf.set_region_center(cv);

    let mut cv = CompositVec::new();
    cv.push(Composit::Start(CompositTag::new("text:p")));
    cv.push(Composit::Text("fimfim".to_string()));
    cv.push(Composit::End("text:p".to_string()));
    hf.set_region_left(cv);

    pl.set_header(hf);

    wb.add_pagelayout(pl);

    write_ods(&wb, "test_out/hf0.ods")?;

    Ok(())
}