//! Text is stored as a simple String whenever possible.
//! When there is a more complex structure, a TextTag is constructed
//! which mirrors the Xml tree structure.

use crate::xmltree::{XmlContent, XmlTag};

pub type TextTag = XmlTag;
pub type TextContent = XmlContent;
