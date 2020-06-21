use spreadsheet_ods::OdsError;
use spreadsheet_ods::refs::{CellRange, CellRef, colname, parse_cellrange, parse_cellranges, parse_cellref, parse_colname, parse_rowname, push_cellrange, push_cellref, push_colname, push_rowname, push_tablename, rowname};

#[test]
fn test_names() {
    let mut buf = String::new();

    push_colname(&mut buf, 0);
    assert_eq!(buf, "A");
    buf.clear();

    push_colname(&mut buf, 1);
    assert_eq!(buf, "B");
    buf.clear();

    push_colname(&mut buf, 26);
    assert_eq!(buf, "AA");
    buf.clear();

    push_colname(&mut buf, 675);
    assert_eq!(buf, "YZ");
    buf.clear();

    push_colname(&mut buf, 676);
    assert_eq!(buf, "ZA");
    buf.clear();

    push_rowname(&mut buf, 0);
    assert_eq!(buf, "1");
    buf.clear();

    push_rowname(&mut buf, 927);
    assert_eq!(buf, "928");
    buf.clear();

    push_tablename(&mut buf, Some(&"fable".to_string()));
    assert_eq!(buf, "fable.");
    buf.clear();

    push_tablename(&mut buf, Some(&"fa le".to_string()));
    assert_eq!(buf, "'fa le'.");
    buf.clear();

    push_tablename(&mut buf, Some(&"fa'le".to_string()));
    assert_eq!(buf, "'fa''le'.");
    buf.clear();

    push_tablename(&mut buf, None);
    assert_eq!(buf, ".");
    buf.clear();

    push_cellref(&mut buf, &CellRef::simple(5, 6));
    assert_eq!(buf, ".G6");
    buf.clear();

    push_cellrange(&mut buf, &CellRange::simple(5, 6, 7, 8));
    assert_eq!(buf, ".G6:.I8");
    buf.clear();

    push_cellrange(&mut buf, &CellRange::table("blame", 5, 6, 7, 8));
    assert_eq!(buf, "blame.G6:blame.I8");
    buf.clear();
}


#[test]
fn test_parse() -> Result<(), OdsError> {
    for i in 0..704 {
        let mut pos = 0usize;
        let cn = colname(i);
        let ccc = parse_colname(&cn, &mut pos);
        assert_eq!(Some(i), ccc);
        assert_eq!(cn.len(), pos);
    }

    for i in 0..101 {
        let mut pos = 0usize;
        let cn = rowname(i);
        let cr = parse_rowname(&cn, &mut pos);
        assert_eq!(Some(i), cr);
        assert_eq!(cn.len(), pos);
    }

    let mut pos = 0usize;
    let cn = "A32";
    let cc = parse_colname(cn, &mut pos);
    assert_eq!(Some(0), cc);
    assert_eq!(1, pos);

    let mut pos = 0usize;
    let cn = "AAAA32 ";
    let cc = parse_colname(cn, &mut pos);
    assert_eq!(Some(18278), cc);
    assert_eq!(4, pos);
    let cr = parse_rowname(cn, &mut pos);
    assert_eq!(Some(31), cr);
    assert_eq!(6, pos);


    let mut pos = 0usize;
    let cn = ".A3";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef::simple(2, 0));

    let mut pos = 0usize;
    let cn = ".$A3";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef {
        table: None,
        row: 2,
        abs_row: false,
        col: 0,
        abs_col: true,
    });

    let mut pos = 0usize;
    let cn = ".A$3";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef {
        table: None,
        row: 2,
        abs_row: true,
        col: 0,
        abs_col: false,
    });

    let mut pos = 0usize;
    let cn = "fufufu.A3";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef {
        table: Some("fufufu".to_string()),
        row: 2,
        abs_row: false,
        col: 0,
        abs_col: false,
    });

    let mut pos = 0usize;
    let cn = "'lak.moi'.A3";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef {
        table: Some("lak.moi".to_string()),
        row: 2,
        abs_row: false,
        col: 0,
        abs_col: false,
    });

    let mut pos = 0usize;
    let cn = "'lak''moi'.A3";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef {
        table: Some("lak'moi".to_string()),
        row: 2,
        abs_row: false,
        col: 0,
        abs_col: false,
    });

    let mut pos = 4usize;
    let cn = "****.B4";
    let cr = parse_cellref(cn, &mut pos)?;
    assert_eq!(cr, CellRef {
        table: None,
        row: 3,
        abs_row: false,
        col: 1,
        abs_col: false,
    });


    let mut pos = 0usize;
    let cn = ".A3:.F9";
    let cr = parse_cellrange(cn, &mut pos)?;
    assert_eq!(cr, CellRange {
        from: CellRef {
            table: None,
            row: 2,
            col: 0,
            abs_row: false,
            abs_col: false,
        },
        to: CellRef {
            table: None,
            row: 8,
            col: 5,
            abs_row: false,
            abs_col: false,
        },
    });

    let mut pos = 0usize;
    let cn = "table.A3:.F9";
    let cr = parse_cellrange(cn, &mut pos)?;
    assert_eq!(cr, CellRange {
        from: CellRef {
            table: Some("table".to_string()),
            row: 2,
            col: 0,
            abs_row: false,
            abs_col: false,
        },
        to: CellRef {
            table: None,
            row: 8,
            col: 5,
            abs_row: false,
            abs_col: false,
        },
    });

    let mut pos = 0usize;
    let cn = "table.A3:fable.F9";
    let cr = parse_cellrange(cn, &mut pos)?;
    assert_eq!(cr, CellRange {
        from: CellRef {
            table: Some("table".to_string()),
            row: 2,
            col: 0,
            abs_row: false,
            abs_col: false,
        },
        to: CellRef {
            table: Some("fable".to_string()),
            row: 8,
            col: 5,
            abs_row: false,
            abs_col: false,
        },
    });

    let mut pos = 0usize;
    let cn = "table.A3:fable.F9 table.A4:fable.F10";
    let cr = parse_cellranges(cn, &mut pos)?;
    assert_eq!(cr, Some(vec![
        CellRange {
            from: CellRef {
                table: Some("table".to_string()),
                row: 2,
                col: 0,
                abs_row: false,
                abs_col: false,
            },
            to: CellRef {
                table: Some("fable".to_string()),
                row: 8,
                col: 5,
                abs_row: false,
                abs_col: false,
            },
        },
        CellRange {
            from: CellRef {
                table: Some("table".to_string()),
                row: 3,
                col: 0,
                abs_row: false,
                abs_col: false,
            },
            to: CellRef {
                table: Some("fable".to_string()),
                row: 9,
                col: 5,
                abs_row: false,
                abs_col: false,
            },
        }
    ])
    );


    Ok(())
}
