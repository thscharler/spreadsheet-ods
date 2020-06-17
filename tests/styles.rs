use color::Rgb;

use spreadsheet_ods::attrmap::{AttrFoBorder, AttrTableCell, Border, WrapOption};
use spreadsheet_ods::style::Style;
use spreadsheet_ods::WorkBook;

#[test]
fn teststyles() {
    let mut wb = WorkBook::new();

    let mut st = Style::cell_style("ce12", "num2");
    st.cell_mut().set_border("0.5pt", Border::Dotted, Rgb::new(0, 0, 0));
    st.cell_mut().set_wrap_option(WrapOption::Wrap);

    wb.add_style(st);
}