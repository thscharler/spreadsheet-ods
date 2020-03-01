use std::collections::HashMap;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

pub fn start(tag: &str) -> Event {
    let b = BytesStart::owned_name(tag.as_bytes());
    Event::Start(b)
}

pub fn start_a<'a>(tag: &'a str, attr: Vec<(&'a str, String)>) -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr {
        b.push_attribute((a, v.as_ref()));
    }

    Event::Start(b)
}

pub fn start_o<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                attr: Option<&'a HashMap<String, String, S>>)
                                                -> Event::<'a> {
    if let Some(attr) = attr {
        start_m(tag, attr)
    } else {
        start(tag)
    }
}

pub fn start_m<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                attr: &'a HashMap<String, String, S>)
                                                -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr {
        b.push_attribute((a.as_str(), v.as_str()));
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
    let b = BytesStart::owned_name(tag.as_bytes());
    Event::Empty(b)
}

pub fn empty_a<'a>(tag: &'a str,
                   attr: Vec<(&'a str, String)>)
                   -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr {
        b.push_attribute((a, v.as_ref()));
    }

    Event::Empty(b)
}

pub fn empty_o<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                attr: Option<&'a HashMap<String, String, S>>)
                                                -> Event::<'a> {
    if let Some(attr) = attr {
        empty_m(tag, attr)
    } else {
        empty(tag)
    }
}

pub fn empty_m<'a, S: ::std::hash::BuildHasher>(tag: &'a str,
                                                attr: &'a HashMap<String, String, S>)
                                                -> Event::<'a> {
    let mut b = BytesStart::owned_name(tag.as_bytes());

    for (a, v) in attr.iter() {
        b.push_attribute((a.as_str(), v.as_str()));
    }

    Event::Empty(b)
}
