use spreadsheet_ods::refs::{push_colname, push_rowname};

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
    assert_eq!(buf, "0");
    buf.clear();

    push_rowname(&mut buf, 927);
    assert_eq!(buf, "927");
    buf.clear();
}