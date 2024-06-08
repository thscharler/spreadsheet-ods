use chrono::{NaiveDate, NaiveDateTime};
use icu_locid::locale;
use spreadsheet_ods::{write_ods, Sheet, WorkBook};
use std::rc::Rc;
use std::time::Duration;

struct Reel {
    url: String,
    ratio: Option<Rc<u32>>,
    account: String,
    like: usize,
    comments: usize,
    views: Option<usize>,
    duration: Duration,
    paid_partnership: bool,
    date: NaiveDateTime,
    caption: String,
}

#[test]
fn test_issue_54() {
    let mut wb = WorkBook::new(locale!("de_AT"));
    let mut sh = Sheet::new("one");

    let mut reels = Vec::new();
    reels.push(Reel {
        url: "https://github.com/thscharler/spreadsheet-ods/issues/54".to_string(),
        ratio: Some(Rc::new(5)),
        account: "nobody".to_string(),
        like: 1,
        comments: 1,
        views: Some(200),
        duration: Duration::default(),
        paid_partnership: false,
        date: NaiveDateTime::default(),
        caption: "something".to_string(),
    });

    for (i, reel) in reels.into_iter().enumerate() {
        dbg!(reel.date.format("%d-%m-%Y %H:%M:%S").to_string());

        let i = i as u32 + 1;
        let url = &reel.url;
        let formula = format!(r#"=HYPERLINK("{url}";"url")"#);
        sh.set_formula(i, 0, formula);
        sh.set_col_width(0, spreadsheet_ods::Length::In(0.40));
        sh.set_value(i, 1, *reel.ratio.unwrap_or_default());
        sh.set_value(i, 2, &reel.account);
        sh.set_col_width(2, spreadsheet_ods::Length::In(1.));
        sh.set_value(i, 3, reel.like as u32);
        sh.set_value(i, 4, reel.comments as u32);
        sh.set_col_width(4, spreadsheet_ods::Length::In(0.7));
        sh.set_value(i, 5, reel.views.unwrap_or_default() as u32);
        sh.set_value(i, 6, reel.duration.as_secs().to_string());
        sh.set_col_width(6, spreadsheet_ods::Length::In(0.6));
        sh.set_value(i, 7, reel.paid_partnership);
        sh.set_col_width(7, spreadsheet_ods::Length::In(0.5));
        sh.set_value(i, 8, &reel.date.format("%d-%m-%Y %H:%M:%S").to_string());
        sh.set_col_width(8, spreadsheet_ods::Length::In(1.35));
        sh.set_value(i, 9, &reel.caption);
        sh.set_col_width(9, spreadsheet_ods::Length::In(15.));
    }

    wb.push_sheet(sh);

    _ = dbg!(write_ods(&mut wb, "test_54.ods"));
}
