use spreadsheet_ods::WorkBook;

#[test]
fn test_default() {
    let wb = WorkBook::default();
    // should not panic, is good
}
