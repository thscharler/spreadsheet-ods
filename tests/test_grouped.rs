use spreadsheet_ods::grouped::{ColGroup, RowGroup};
use spreadsheet_ods::{read_ods, write_ods, OdsError, Sheet, WorkBook};

#[test]
fn test_write_group1() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(0, 0, 1);
    sh.set_value(1, 0, 1);
    sh.set_value(2, 0, 1);
    sh.set_value(3, 0, 1);
    sh.set_value(4, 0, 1);
    sh.set_value(5, 0, 1);
    sh.set_value(6, 0, 1);
    sh.set_value(7, 0, 1);
    sh.set_value(8, 0, 1);
    sh.set_value(9, 0, 1);

    sh.add_row_group(RowGroup::new(1, 4, true));
    sh.add_row_group(RowGroup::new(1, 2, true));
    sh.add_row_group(RowGroup::new(1, 3, true));
    sh.add_row_group(RowGroup::new(4, 4, true));
    //sh.add_row_group(RowGroup::new(4, 5, true));
    sh.add_row_group(RowGroup::new(6, 9, false));
    sh.add_row_group(RowGroup::new(7, 9, true));
    sh.add_row_group(RowGroup::new(8, 9, true));
    sh.add_row_group(RowGroup::new(9, 9, true));
    sh.add_row_group(RowGroup::new(40, 45, true));

    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/rowgroup1.ods")?;

    let wb = read_ods("test_out/rowgroup1.ods")?;
    let sh = wb.sheet(0);

    let v = sh.row_group_iter().cloned().collect::<Vec<_>>();
    assert!(v.contains(&RowGroup::new(1, 4, true)));
    assert!(v.contains(&RowGroup::new(1, 2, true)));
    assert!(v.contains(&RowGroup::new(1, 3, true)));
    assert!(v.contains(&RowGroup::new(4, 4, true)));
    assert!(v.contains(&RowGroup::new(6, 9, false)));
    assert!(v.contains(&RowGroup::new(7, 9, true)));
    assert!(v.contains(&RowGroup::new(8, 9, true)));
    assert!(v.contains(&RowGroup::new(9, 9, true)));

    Ok(())
}

#[test]
fn test_write_group2() -> Result<(), OdsError> {
    let mut wb = WorkBook::new_empty();

    let mut sh = Sheet::new("Sheet1");
    sh.set_value(0, 0, 1);
    sh.set_value(1, 1, 1);
    sh.set_value(2, 2, 1);
    sh.set_value(3, 3, 1);
    sh.set_value(4, 4, 1);
    sh.set_value(5, 5, 1);
    sh.set_value(6, 6, 1);
    sh.set_value(7, 7, 1);
    sh.set_value(8, 8, 1);
    sh.set_value(9, 9, 1);

    sh.add_col_group(ColGroup::new(1, 4, true));
    sh.add_col_group(ColGroup::new(1, 2, true));
    sh.add_col_group(ColGroup::new(1, 3, true));
    sh.add_col_group(ColGroup::new(4, 4, true));
    //sh.add_col_group(ColGroup::new(4, 5, true));
    sh.add_col_group(ColGroup::new(6, 9, false));
    sh.add_col_group(ColGroup::new(7, 9, true));
    sh.add_col_group(ColGroup::new(8, 9, true));
    sh.add_col_group(ColGroup::new(9, 9, true));
    sh.add_col_group(ColGroup::new(40, 45, true));

    wb.push_sheet(sh);

    write_ods(&mut wb, "test_out/colgroup2.ods")?;

    let wb = read_ods("test_out/colgroup2.ods")?;
    let sh = wb.sheet(0);

    dbg!(sh);

    let v = sh.col_group_iter().cloned().collect::<Vec<_>>();
    assert!(v.contains(&ColGroup::new(1, 4, true)));
    assert!(v.contains(&ColGroup::new(1, 2, true)));
    assert!(v.contains(&ColGroup::new(1, 3, true)));
    assert!(v.contains(&ColGroup::new(4, 4, true)));
    assert!(v.contains(&ColGroup::new(6, 9, false)));
    assert!(v.contains(&ColGroup::new(7, 9, true)));
    assert!(v.contains(&ColGroup::new(8, 9, true)));
    assert!(v.contains(&ColGroup::new(9, 9, true)));

    Ok(())
}
