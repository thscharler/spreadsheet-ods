use color::Rgb;

use spreadsheet_ods::WorkBook;
use spreadsheet_ods::attrmap::{AttrFoBorder, AttrTableCell, Border, WrapOption};
use spreadsheet_ods::style::Style;

#[test]
fn teststyles() {
    let mut wb = WorkBook::new();

    let mut st = Style::cell_style("ce12", "num2");
    st.cell_attr().set_border("0.5pt", Border::Dotted, Rgb::new(0, 0, 0));
    st.cell_attr().set_wrap_option(WrapOption::Wrap);

    wb.add_style(st);
}