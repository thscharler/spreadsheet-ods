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
