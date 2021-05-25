use spreadsheet_ods::{
    currency, percent, read_ods, write_ods, OdsError, SCell, Sheet, Value, ValueType, WorkBook,
};

#[test]
fn test_workbook() {
    let mut wb = WorkBook::new();

    let sh = Sheet::new();
    // println!("sizeof Sheet {}", size_of_val(&sh));
    wb.push_sheet(sh);
    assert_eq!(wb.num_sheets(), 1);
    wb.push_sheet(Sheet::new_with_name("b"));
    wb.push_sheet(Sheet::new_with_name("c"));
    assert_eq!(wb.sheet(1).name(), "b");
    wb.insert_sheet(1, Sheet::new_with_name("x"));
    assert_eq!(wb.sheet(1).name(), "x");
    let sh = wb.remove_sheet(1);
    assert_eq!(sh.name(), "x");
    assert_eq!(wb.num_sheets(), 3);
}

#[test]
fn test_def_style() {
    let mut wb = WorkBook::new();

    wb.add_def_style(ValueType::Number, &"val0".into());
    assert_eq!(wb.def_style(ValueType::Number), Some(&"val0".to_string()));
    assert!(wb.def_style(ValueType::Text).is_none());
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
fn test_macros() {
    let mut sh = Sheet::new();

    sh.set_value(0, 0, currency!("€", 20));
    sh.set_value(0, 0, percent!(17.22));
}
