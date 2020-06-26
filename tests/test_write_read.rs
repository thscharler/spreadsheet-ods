use spreadsheet_ods::{OdsError, read_ods, Sheet, ValueType, WorkBook, write_ods};
use spreadsheet_ods::refs::{CellRange, ColRange, RowRange};

#[test]
fn test_0() -> Result<(), OdsError> {
    if cfg!(dump_xml) {
        println!("dump_xml!");
        panic!();
    } else {
        println!("no dump_xml!");
    }


    println!("test_0");
    let mut wb = WorkBook::new();
    let mut sh = Sheet::new();

    sh.set_value(0, 0, "A");

    wb.push_sheet(sh);

    write_ods(&wb, "test_out/test_0.ods")?;

    let wi = read_ods("test_out/test_0.ods")?;
    let si = wi.sheet(0);

    println!("{:?}", si);


    assert_eq!(si.value(0, 0).as_str_or(""), "A");

    Ok(())
}

#[test]
fn test_span() -> Result<(), OdsError> {
    println!("test_span");

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

    write_ods(&wb, "test_out/test_span.ods")?;
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
    println!("test_header");
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

    write_ods(&wb, "test_out/test_header0.ods")?;

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
    println!("test_print_range");
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


    write_ods(&wb, "test_out/test_print_range.ods")?;

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
fn read_text() -> Result<(), OdsError> {
    println!("read_text");

    let wb = read_ods("tests/text.ods")?;
    let sh = wb.sheet(0);

    let v = sh.value(0, 0);

    assert_eq!(v.value_type(), ValueType::TextXml);

    println!("text value {:?}", v);

    Ok(())
}

#[test]
fn read_orders() -> Result<(), OdsError> {
    println!("read_orders");

    let _wb = read_ods("tests/orders.ods");

    Ok(())
}