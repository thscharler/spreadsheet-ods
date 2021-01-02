//! Defines an XML-Tree. This is used for parts of the spreadsheet
//! that are not destructured in detail, but simply passed through.
//! With a little bit of luck there is still some meaning left after
//! modifying the rest.
//!
//! ```
//! use spreadsheet_ods::xmltree::XmlTag;
//! use spreadsheet_ods::xmltree::AttrMap2Trait;
//!
//! let tag = XmlTag::new("table:shapes")
//!         .con_tag(XmlTag::new("draw:frame")
//!             .con_attr("draw:z", "0")
//!             .con_attr("draw:name", "Bild 1")
//!             .con_attr("draw:style:name", "gr1")
//!             .con_attr("draw:text-style-name", "P1")
//!             .con_attr("svg:width", "10.198cm")
//!             .con_attr("svg:height", "1.75cm")
//!             .con_attr("svg:x", "0cm")
//!             .con_attr("svg:y", "0cm")
//!             .con_tag(XmlTag::new("draw:image")
//!                 .con_attr("xlink:href", "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg")
//!                 .con_attr("xlink:type", "simple")
//!                 .con_attr("xlink:show", "embed")
//!                 .con_attr("xlink:actuate", "onLoad")
//!                 .con_attr("loext:mime-type", "image/jpeg")
//!                 .con_tag(XmlTag::new("text:p")
//!                     .con_text("sometext")
//!                 )
//!             )
//!         );
//!
//! // or
//! let mut tag = XmlTag::new("table:shapes");
//! tag.set_attr("draw:z", "0".to_string());
//! tag.set_attr("draw:name", "Bild 1".to_string());
//! tag.set_attr("draw:style:name", "gr1".to_string());
//!
//! let mut tag2 = XmlTag::new("draw:image");
//! tag2.set_attr("xlink:type", "simple".to_string());
//! tag2.set_attr("xlink:show", "embed".to_string());
//! tag2.push_text("some text");
//! tag.push_tag(tag2);
//!
//! ```

pub use crate::attrmap2::{AttrMap2, AttrMap2Trait};

/// Defines a XML tag and it's children.
#[derive(Debug, Clone, Default)]
pub struct XmlTag {
    name: String,
    attr: AttrMap2,
    content: Vec<XmlContent>,
}

impl AttrMap2Trait for XmlTag {
    fn attr_map(&self) -> &AttrMap2 {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }
}

impl XmlTag {
    /// New Tag.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            attr: Default::default(),
            content: vec![],
        }
    }

    /// Name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Any text or child elements?
    pub fn is_empty(&self) -> bool {
        self.content.is_empty()
    }

    /// Add an element.
    pub fn push_tag(&mut self, xmltag: XmlTag) {
        self.content.push(XmlContent::Tag(xmltag));
    }

    /// Retrieves the first tag, if any.
    pub fn pop_tag(&mut self) -> Option<XmlTag> {
        match self.content.get(0) {
            None => None,
            Some(XmlContent::Tag(_)) => {
                if let XmlContent::Tag(tag) = self.content.pop().unwrap() {
                    Some(tag)
                } else {
                    unreachable!()
                }
            }
            Some(XmlContent::Text(_)) => None,
        }
    }

    /// Add text.
    pub fn push_text<S: Into<String>>(&mut self, text: S) {
        self.content.push(XmlContent::Text(text.into()));
    }

    /// Retrieves the first text element.
    pub fn pop_text(&mut self) -> Option<String> {
        match self.content.get(0) {
            None => None,
            Some(XmlContent::Text(_)) => {
                if let XmlContent::Text(text) = self.content.pop().unwrap() {
                    Some(text)
                } else {
                    unreachable!()
                }
            }
            Some(XmlContent::Tag(_)) => None,
        }
    }

    /// Sets an attribute. Allows for cascading.
    pub fn con_attr<'a, S0, S1>(mut self, name: S0, value: S1) -> Self
    where
        S0: Into<&'a str>,
        S1: Into<String>,
    {
        self.attr.set_attr(name.into(), value.into());
        self
    }

    /// Adds an element. Allows for cascading.
    pub fn con_tag(mut self, xmltag: XmlTag) -> Self {
        self.content.push(XmlContent::Tag(xmltag));
        self
    }

    /// Adds text. Allows for cascading.
    pub fn con_text<S: Into<String>>(mut self, text: S) -> Self {
        self.content.push(XmlContent::Text(text.into()));
        self
    }

    /// Returns the content vec.
    pub fn content(&self) -> &Vec<XmlContent> {
        &self.content
    }
}

/// Values of the content vec.
#[derive(Debug, Clone)]
pub enum XmlContent {
    Text(String),
    Tag(XmlTag),
}
