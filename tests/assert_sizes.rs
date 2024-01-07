use spreadsheet_ods::CellStyleRef;
use std::mem::size_of;

#[test]
fn assert_sizes() {
    assert_eq!(size_of::<Option<CellStyleRef>>(), size_of::<CellStyleRef>());
}
