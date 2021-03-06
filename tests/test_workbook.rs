use spreadsheet_ods::{Sheet, ValueType, WorkBook};

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
