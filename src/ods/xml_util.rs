use std::collections::HashMap;
use string_cache::DefaultAtom;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

pub fn start(tag: &str) -> Event {
    let b = BytesStart::borrowed_name(tag.as_bytes());
    Event::Start(b)
}

pub fn start_a<'a>(tag: &'a str, attr: &[(&'a str, &'a str)]) -> Event::<'a> {
    let mut b = BytesStart::borrowed_name(tag.as_bytes());
    for av in attr {
        b.push_attribute(*av);
    }
    Event::Start(b)
}

pub fn start_opt<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                  attr: Option<&'a HashMap<DefaultAtom, String, S>>)
                                                  -> Event::<'a> {
    if let Some(attr) = attr {
        start_m(tag, attr)
    } else {
        start(tag)
    }
}

pub fn start_m<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                attr: &'a HashMap<DefaultAtom, String, S>)
                                                -> Event::<'a> {
    let mut b = BytesStart::borrowed_name(tag.as_bytes());
    for (a, v) in attr {
        b.push_attribute((a.as_ref(), v.as_str()));
    }
    Event::Start(b)
}

pub fn start_am<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                 attr0: &[(&'a str, &'a str)],
                                                 attr1: Option<&'a HashMap<DefaultAtom, String, S>>)
                                                 -> Event::<'a> {
    let mut b = BytesStart::borrowed_name(tag.as_bytes());
    for av in attr0 {
        b.push_attribute(*av);
    }
    if let Some(attr1) = attr1 {
        for (a, v) in attr1 {
            b.push_attribute((a.as_ref(), v.as_str()));
        }
    }
    Event::Start(b)
}

pub fn text(text: &str) -> Event {
    Event::Text(BytesText::from_plain_str(text))
}

pub fn end(tag: &str) -> Event {
    let b = BytesEnd::borrowed(tag.as_bytes());
    Event::End(b)
}

pub fn empty(tag: &str) -> Event {
    let b = BytesStart::borrowed_name(tag.as_bytes());
    Event::Empty(b)
}

pub fn empty_a<'a>(tag: &'a str,
                   attr: &[(&'a str, &'a str)])
                   -> Event::<'a> {
    let mut b = BytesStart::borrowed_name(tag.as_bytes());
    for av in attr {
        b.push_attribute(*av);
    }
    Event::Empty(b)
}

pub fn empty_opt<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                  attr: Option<&'a HashMap<DefaultAtom, String, S>>)
                                                  -> Event::<'a> {
    if let Some(attr) = attr {
        empty_m(tag, attr)
    } else {
        empty(tag)
    }
}

pub fn empty_am<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                 attr0: &[(&'a str, &'a str)],
                                                 attr1: Option<&'a HashMap<DefaultAtom, String, S>>)
                                                 -> Event::<'a> {
    let mut b = BytesStart::borrowed_name(tag.as_bytes());
    for av in attr0 {
        b.push_attribute(*av);
    }
    if let Some(attr1) = attr1 {
        for (a, v) in attr1.iter() {
            b.push_attribute((a.as_ref(), v.as_str()));
        }
    }
    Event::Empty(b)
}

pub fn empty_m<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                attr: &'a HashMap<DefaultAtom, String, S>)
                                                -> Event::<'a> {
    let mut b = BytesStart::borrowed_name(tag.as_bytes());
    for (a, v) in attr.iter() {
        b.push_attribute((a.as_ref(), v.as_str()));
    }
    Event::Empty(b)
}
