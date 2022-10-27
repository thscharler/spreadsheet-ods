use crate::attrmap2::AttrMap2;
use crate::style::{StyleOrigin, StyleUse};
use std::fmt::{Display, Formatter};

style_ref!(GraphicStyleRef);

/// Styles of this type can occur in an odt file.
/// This is only used as a place to put this stuff when reading the ods.
///
#[derive(Debug, Clone)]
pub struct GraphicStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// Style name
    name: String,
    /// General attributes
    // ok style:auto-update 19.467 => ALL
    // ok style:class 19.470, => ALL
    // ignore style:data-style-name 19.473, => CELL, CHART
    // ignore style:default-outlinelevel 19.474, => PARAGRAPH
    // ok style:display-name 19.476, => ALL
    // ignore style:family 19.480, => Not mapped as an attribute.
    // ignore style:list-level 19.499, => PARAGRAPH
    // ignore style:list-style-name 19.500, => PARAGRAPH
    // ignore style:master-page-name 19.501, => PARAGRAPH, TABLE
    // ignore style:name 19.502, => Not mapped as an attribute.
    // ignore style:next-style-name 19.503, => PARAGRAPH
    // ok style:parent-style-name 19.510 => ALL
    // ignore style:percentage-data-style-name 19.511. => PARAGRAPH?
    attr: AttrMap2,
    /// Table style properties
    // ignore these attributes for now.
    graphicstyle: AttrMap2,
}

styles_styles!(GraphicStyle, GraphicStyleRef);

impl GraphicStyle {
    /// Empty.
    pub fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            graphicstyle: Default::default(),
        }
    }

    /// New graphic style.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            graphicstyle: Default::default(),
        }
    }

    /// General attributes.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Graphic style attributes.
    pub fn graphicstyle(&self) -> &AttrMap2 {
        &self.graphicstyle
    }

    /// Graphic style attributes.
    pub fn graphicstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.graphicstyle
    }
}
