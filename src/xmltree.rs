//! Defines an XML-Tree. This is used for parts of the spreadsheet
//! that are not destructured in detail, but simply passed through.
//! With a little bit of luck there is still some meaning left after
//! modifying the rest.
//!
//! ```
//! use spreadsheet_ods::xmltree::XmlTag;
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
//!                 .attr_slice(&[
//!                     ("xlink:href", "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg".into()),
//!                     ("xlink:type", "simple".into()),
//!                     ("xlink:show", "embed".into()),
//!                     ("xlink:actuate", "onLoad".into()),
//!                     ("loext:mime-type", "image/jpeg".into()),
//!                 ])
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

use crate::attrmap2::AttrMap2;
use crate::text::TextP;
use crate::OdsError;
use std::fmt::{Display, Formatter};

/// Defines a XML tag and it's children.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct XmlTag {
    name: String,
    attr: AttrMap2,
    content: Vec<XmlContent>,
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

/// Functionality for vectors of XmlTag's.
pub trait XmlVec {
    /// Adds the text as new XmlTag of text:p.
    fn add_text<S: Into<String>>(&mut self, txt: S);

    /// Adds the tag as new XmlTag.
    fn add_tag<T: Into<XmlTag>>(&mut self, tag: T);
}

impl XmlVec for &mut Vec<XmlTag> {
    fn add_text<S: Into<String>>(&mut self, txt: S) {
        self.push(TextP::new().text(txt).into());
    }

    fn add_tag<T: Into<XmlTag>>(&mut self, tag: T) {
        self.push(tag.into());
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

    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Sets an attribute
    pub fn set_attr<'a, S: Into<&'a str>, T: Into<String>>(&mut self, name: S, value: T) {
        self.attr.set_attr(name.into(), value.into());
    }

    /// Gets an attribute
    pub fn get_attr<'a, S: Into<&'a str>>(&self, name: S) -> Option<&String> {
        self.attr.attr(name.into())
    }

    /// Adds more attributes.
    pub fn add_attr_slice(&mut self, attr: &[(&str, String)]) {
        self.attr.add_all(attr);
    }

    /// Add an element.
    pub fn add_tag<T: Into<XmlTag>>(&mut self, tag: T) {
        self.content.push(XmlContent::Tag(tag.into()));
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

    /// Extracts the plain text from this tag and its content.
    pub fn extract_text(&self, buf: &mut String) {
        for c in &self.content {
            match c {
                XmlContent::Text(t) => {
                    buf.push_str(t.as_str());
                }
                XmlContent::Tag(t) => {
                    t.extract_text(buf);
                }
            }
        }
    }

    /// Converts the content into a `Vec<XmlTag>`. Any occurring text content
    /// is an error.
    pub fn into_vec(self) -> Result<Vec<XmlTag>, OdsError> {
        let mut content = Vec::new();

        for c in self.content {
            match c {
                XmlContent::Text(v) => {
                    return Err(OdsError::Parse("Unexpected literal text ", Some(v)))
                }
                XmlContent::Tag(v) => content.push(v),
            }
        }

        Ok(content)
    }

    /// Converts the content into a `Vec<XmlTag>`. Any occurring text content
    /// is ok.
    pub fn into_mixed_vec(self) -> Vec<XmlContent> {
        self.content
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

///
/// A XmlTag can contain any mixture of XmlTags and text content.
///
#[derive(Debug, Clone, PartialEq)]
#[allow(variant_size_differences)]
pub enum XmlContent {
    /// Text content.
    Text(String),
    /// Contained xml-tags.
    Tag(XmlTag),
}
