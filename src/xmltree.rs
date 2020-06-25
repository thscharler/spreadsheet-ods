//! Defines an XML-Tree. This is used for parts of the spreadsheet
//! that are not destructured in detail, but simply passed through.
//! With a little bit of luck there is still some meaning left after
//! modifying the rest.

use crate::attrmap::{AttrMapType, AttrMap, AttrMapIter};
use std::collections::HashMap;
use string_cache::DefaultAtom;

#[derive(Debug, Clone, Default)]
pub struct XmlTag {
    name: String,
    empty: bool,
    attr: AttrMapType,
    content: Vec<XmlContent>,
}

impl AttrMap for XmlTag {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl XmlTag {
    pub fn new<S: Into<String>>(name: S, empty: bool) -> Self {
        Self {
            name: name.into(),
            empty,
            attr: None,
            content: vec![],
        }
    }

    pub fn tag<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            empty: false,
            attr: None,
            content: vec![],
        }
    }

    pub fn empty<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            empty: true,
            attr: None,
            content: vec![],
        }
    }

    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn set_empty(&mut self, empty: bool) {
        self.empty = empty;
    }

    pub fn is_empty(&self) -> bool {
        self.empty
    }

    pub fn attr_iter(&self) -> AttrMapIter {
        AttrMapIter::from(self.attr_map())
    }

    pub fn add_tag(&mut self, xmltag: XmlTag) {
        self.content.push(XmlContent::Tag(xmltag));
    }

    pub fn add_text<S: Into<String>>(&mut self, text: S) {
        self.content.push(XmlContent::Text(text.into()));
    }

    pub fn attr_con<'a, S0, S1>(mut self, name: S0, value: S1) -> Self
        where S0: Into<&'a str>,
              S1: Into<String>
    {
        self.attr_map_mut()
            .get_or_insert_with(|| Box::new(HashMap::new()))
            .insert(DefaultAtom::from(name.into()), value.into());
        self
    }

    pub fn tag_con(mut self, xmltag: XmlTag) -> Self {
        self.content.push(XmlContent::Tag(xmltag));
        self
    }

    pub fn text_con<S: Into<String>>(mut self, text: S) -> Self {
        self.content.push(XmlContent::Text(text.into()));
        self
    }

    pub fn content(&self) -> &Vec<XmlContent> {
        &self.content
    }
}

#[derive(Debug, Clone)]
pub enum XmlContent {
    Text(String),
    Tag(XmlTag),
}






















