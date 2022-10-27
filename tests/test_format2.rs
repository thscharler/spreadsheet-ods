use icu_locid::locale;
use spreadsheet_ods::format::FormatCalendarStyle;
use spreadsheet_ods::{format, OdsError, ValueFormat, ValueType};

// #[test]
// fn read_orders() -> Result<(), OdsError> {
//     let wb = read_ods("tests/Unbenannt 1.ods")?;
//
//     dbg!(wb);
//
//     Ok(())
// }

#[test]
fn builder() -> Result<(), OdsError> {
    let mut vf = ValueFormat::new_localized("fro", locale!("de_AT"), ValueType::Percentage);

    format::PartDayBuilder::new(&mut vf)
        .calendar(FormatCalendarStyle::Gregorian)
        .short_style()
        .build();

    Ok(())
}
