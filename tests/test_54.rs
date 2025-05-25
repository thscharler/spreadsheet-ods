use chrono::NaiveDateTime;
use spreadsheet_ods::{write_ods, Sheet, WorkBook};
use std::sync::mpsc;
use std::time::Duration;

#[derive(Default)]
struct Reel {
    url: String,
    ratio: Option<i8>,
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
    let (tx, rx) = mpsc::channel();
    let handle = std::thread::spawn(move || {
        let mut wb = WorkBook::new_empty();
        let mut sh = Sheet::new("one");

        let mut reels = Vec::new();
        // works with 1, but not with more
        (0..10).for_each(|i| {
            reels.push(Reel {
                caption: "i".repeat(i),
                ..Default::default()
            });
        });

        println!("data is {:?} reels", reels.len());
        for (i, reel) in reels.into_iter().enumerate() {
            let i = i as u32 + 1;
            sh.set_formula(i, 0, reel.url);
            sh.set_col_width(0, spreadsheet_ods::Length::In(0.40));
            sh.set_value(i, 1, reel.ratio.unwrap_or_default());
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
            sh.set_value(i, 8, reel.date.format("%d-%m-%Y %H:%M:%S").to_string());
            sh.set_col_width(8, spreadsheet_ods::Length::In(1.35));
            sh.set_value(i, 9, &reel.caption);
            sh.set_col_width(9, spreadsheet_ods::Length::In(15.));
        }

        wb.push_sheet(sh);

        dbg!(write_ods(&mut wb, "test_out/test_54.ods").expect("can't write file"));

        tx.send(()).unwrap();
    });

    // 3 seconds should be plenty enough
    match rx.recv_timeout(Duration::from_secs(3)) {
        Ok(_) => handle.join().expect("thread panicked"),
        Err(_) => panic!("timeout reached"),
    };
}
