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
//!         .tag(XmlTag::new("draw:frame")
//!             .attr("draw:z", "0")
//!             .attr("draw:name", "Bild 1")
//!             .attr("draw:style:name", "gr1")
//!             .attr("draw:text-style-name", "P1")
//!             .attr("svg:width", "10.198cm")
//!             .attr("svg:height", "1.75cm")
//!             .attr("svg:x", "0cm")
//!             .attr("svg:y", "0cm")
//!             .tag(XmlTag::new("draw:image")
//!                 .attr("xlink:href", "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg")
//!                 .attr("xlink:type", "simple")
//!                 .attr("xlink:show", "embed")
//!                 .attr("xlink:actuate", "onLoad")
//!                 .attr("loext:mime-type", "image/jpeg")
//!                 .tag(XmlTag::new("text:p")
//!                     .text("sometext")
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
//! tag2.add_text("some text");
//! tag.add_tag(tag2);
//!
//! ```

pub use crate::attrmap2::{AttrMap2, AttrMap2Trait};
use std::fmt::{Display, Formatter};

/// Defines a XML tag and it's children.
#[derive(Debug, Clone, Default)]
pub struct XmlTag {
    name: String,
    attr: AttrMap2,
    content: Vec<XmlContent>,
}

impl AttrMap2Trait for XmlTag {
    fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }
}

impl From<&str> for XmlTag {
    fn from(name: &str) -> Self {
        XmlTag::new(name)
    }
}

impl From<String> for XmlTag {
    fn from(name: String) -> Self {
        XmlTag::new(name)
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

    /// Sets an attribute
    pub fn set_attr<'a, S: Into<&'a str>, T: Into<String>>(&mut self, name: S, value: T) {
        self.attr.set_attr(name.into(), value.into());
    }

    /// Adds more attributes.
    pub fn add_attr_slice(&mut self, attr: &[(&str, String)]) {
        self.attr.add_all(attr);
    }

    /// Add an element.
    pub fn add_tag<T: Into<XmlTag>>(&mut self, tag: T) {
        self.content.push(XmlContent::Tag(tag.into()));
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
    pub fn add_text<S: Into<String>>(&mut self, text: S) {
        self.content.push(XmlContent::Text(text.into()));
    }

    /// Sets an attribute. Allows for cascading.
    pub fn attr<'a, S: Into<&'a str>, T: Into<String>>(mut self, name: S, value: T) -> Self {
        self.set_attr(name, value);
        self
    }

    /// Adds more attributes.
    pub fn attr_slice(mut self, attr: &[(&str, String)]) -> Self {
        self.add_attr_slice(attr);
        self
    }

    /// Adds an element. Allows for cascading.
    pub fn tag<T: Into<XmlTag>>(mut self, tag: T) -> Self {
        self.add_tag(tag);
        self
    }

    /// Adds text. Allows for cascading.
    pub fn text<S: Into<String>>(mut self, text: S) -> Self {
        self.add_text(text);
        self
    }

    /// Returns the content vec.
    pub fn content(&self) -> &Vec<XmlContent> {
        &self.content
    }

    /// Returns the content vec.
    pub fn content_mut(&mut self) -> &mut Vec<XmlContent> {
        &mut self.content
    }
}

impl Display for XmlTag {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        write!(f, "<{}", self.name)?;
        for (n, v) in self.attr.iter() {
            write!(f, " {}=\"{}\"", n, v)?;
        }
        if self.content.is_empty() {
            writeln!(f, "/>")?;
        } else {
            writeln!(f, ">")?;

            for c in &self.content {
                match c {
                    XmlContent::Text(t) => {
                        writeln!(f, "{}", t)?;
                    }
                    XmlContent::Tag(t) => {
                        t.fmt(f)?;
                    }
                }
            }

            writeln!(f, "</{}>", self.name)?;
        }

        Ok(())
    }
}

/// Values of the content vec.
#[derive(Debug, Clone)]
pub enum XmlContent {
    Text(String),
    Tag(XmlTag),
}
