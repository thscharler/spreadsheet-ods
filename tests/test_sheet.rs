use spreadsheet_ods::{
    cm, currency, percent, read_ods, write_ods, CellRange, ColRange, Length, OdsError, RowRange,
    SCell, Sheet, Value, ValueType, WorkBook,
};

#[test]
fn test_colwidth() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut sh = Sheet::new_with_name("Sheet1");
    sh.set_value(0, 0, 1234);
    sh.set_col_width(0, cm!(2.54));
    sh.set_row_height(0, cm!(1.27));
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/colwidth.ods")?;

    let wb = read_ods("test_out/colwidth.ods")?;
    assert_eq!(wb.sheet(0).col_width(0), cm!(2.54));
    assert_eq!(wb.sheet(0).row_height(0), cm!(1.27));

    Ok(())
}

#[test]
fn test_cell() {
    let mut sh = Sheet::new();

    sh.set_value(5, 5, 1);
    sh.set_value(6, 6, 2);

    if let Some(c) = sh.cell(5, 5) {
        assert_eq!(c.value().as_i32_or(0), 1);
    }

    let c = sh.cell_mut(6, 6);
    c.set_value(3);
    let mut x = SCell::new();
    std::mem::swap(c, &mut x);
    assert_eq!(x.value().as_f64_or(0.0), 3.0);
}

#[test]
fn test_row_repeat() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();
    let mut sh = Sheet::new();

    sh.set_value(2, 2, 1);
    sh.set_value(4, 4, 2);
    sh.set_row_repeat(4, 2);

    wb.push_sheet(sh);
    write_ods(&mut wb, "test_out/row_repeat.ods")?;

    let wb = read_ods("test_out/row_repeat.ods")?;
    assert_eq!(wb.sheet(0).row_repeat(4), 2);

    Ok(())
}

#[test]
fn test_value_macros() {
    let mut sh = Sheet::new();

    sh.set_value(0, 0, currency!("€", 20));
    assert_eq!(sh.value(0, 0).value_type(), ValueType::Currency);
    sh.set_value(0, 0, percent!(17.22));
    assert_eq!(sh.value(0, 0).value_type(), ValueType::Percentage);
}

#[test]
fn test_span() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut sh = Sheet::new();
    sh.set_value(0, 0, "A");
    sh.set_value(0, 1, "A2");
    sh.set_value(0, 2, "bomb");
    sh.set_value(1, 0, "bomb");
    sh.set_value(1, 1, "bomb");
    sh.set_value(1, 2, "bomb");
    sh.set_col_span(0, 0, 2);
    wb.push_sheet(sh);

    let mut sh = Sheet::new();
    sh.set_value(1, 0, "B");
    sh.set_value(2, 0, "B2");
    sh.set_value(1, 1, "bomb");
    sh.set_value(2, 1, "bomb");
    sh.set_value(3, 0, "bomb");
    sh.set_value(3, 1, "bomb");
    sh.set_row_span(1, 0, 2);
    wb.push_sheet(sh);

    let mut sh = Sheet::new();
    sh.set_value(3, 0, "C");
    sh.set_value(3, 1, "C2");
    sh.set_value(4, 0, "C2");
    sh.set_value(4, 1, "C2");
    sh.set_value(3, 2, "bomb");
    sh.set_value(4, 2, "bomb");
    sh.set_value(5, 0, "bomb");
    sh.set_value(5, 1, "bomb");
    sh.set_value(5, 2, "bomb");
    sh.set_col_span(3, 0, 2);
    sh.set_row_span(3, 0, 2);

    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/test_span.ods")?;
    let wi = read_ods("test_out/test_span.ods")?;

    let si = wi.sheet(0);

    assert_eq!(si.value(0, 0).as_str_or(""), "A");
    assert_eq!(si.col_span(0, 0), 2);

    let si = wi.sheet(1);

    assert_eq!(si.value(1, 0).as_str_or(""), "B");
    assert_eq!(si.row_span(1, 0), 2);

    Ok(())
}

#[test]
fn test_header() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut sh = Sheet::new();
    for i in 0..10 {
        for j in 0..10 {
            sh.set_value(i, j, i + j);
        }
    }
    sh.set_header_cols(0, 2);
    sh.set_header_rows(0, 2);
    wb.push_sheet(sh);

    let mut sh = Sheet::new();
    sh.set_value(0, 0, 0);
    sh.set_value(9, 0, 0);
    sh.set_header_rows(2, 3);
    wb.push_sheet(sh);

    let mut sh = Sheet::new();
    sh.set_value(0, 0, 0);
    sh.set_value(9, 0, 0);
    sh.set_header_rows(0, 3);
    wb.push_sheet(sh);

    let mut sh = Sheet::new();
    sh.set_value(0, 0, 0);
    sh.set_value(9, 0, 0);
    sh.set_header_rows(2, 9);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/test_header0.ods")?;

    let wb = read_ods("test_out/test_header0.ods")?;

    assert_eq!(wb.sheet(0).header_rows().clone(), Some(RowRange::new(0, 2)));
    assert_eq!(wb.sheet(0).header_cols().clone(), Some(ColRange::new(0, 2)));
    assert_eq!(wb.sheet(1).header_rows().clone(), Some(RowRange::new(2, 3)));
    assert_eq!(wb.sheet(2).header_rows().clone(), Some(RowRange::new(0, 3)));
    assert_eq!(wb.sheet(3).header_rows().clone(), Some(RowRange::new(2, 9)));

    Ok(())
}

#[test]
fn test_print_range() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut sh = Sheet::new();
    for i in 0..10 {
        for j in 0..10 {
            sh.set_value(i, j, i * j);
        }
    }
    sh.set_header_cols(0, 0);
    sh.set_header_rows(0, 0);
    sh.add_print_range(CellRange::local(1, 1, 9, 9));
    sh.add_print_range(CellRange::local(11, 11, 19, 19));
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/test_print_range.ods")?;

    let wb = read_ods("test_out/test_print_range.ods")?;
    let sh = wb.sheet(0);

    assert_eq!(wb.sheet(0).header_rows().clone(), Some(RowRange::new(0, 0)));
    assert_eq!(wb.sheet(0).header_cols().clone(), Some(ColRange::new(0, 0)));

    let r = sh.print_ranges().unwrap();
    assert_eq!(r[0], CellRange::local(1, 1, 9, 9));
    assert_eq!(r[1], CellRange::local(11, 11, 19, 19));

    Ok(())
}

#[test]
fn display_print() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();
    let mut s0 = Sheet::new();
    s0.set_value(0, 0, "display");
    s0.set_display(false);
    wb.push_sheet(s0);

    let mut s1 = Sheet::new();
    s1.set_value(0, 0, "print");
    s1.set_print(false);
    wb.push_sheet(s1);

    write_ods(&mut wb, "test_out/display_print.ods")?;

    Ok(())
}

#[test]
fn split_table() -> Result<(), OdsError> {
    let mut wb = WorkBook::new();

    let mut sh = Sheet::new_with_name("Split0");
    sh.set_value(0, 0, 1);
    sh.set_value(0, 1, 2);
    sh.set_value(1, 0, 3);
    sh.set_value(1, 1, 4);
    sh.split_hor_cell(3);
    wb.push_sheet(sh);

    let mut sh = Sheet::new_with_name("Split1");
    sh.set_value(0, 0, 1);
    sh.set_value(0, 1, 2);
    sh.set_value(1, 0, 3);
    sh.set_value(1, 1, 4);
    sh.split_hor_pixel(250);
    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/split_table.ods")?;

    Ok(())
}
