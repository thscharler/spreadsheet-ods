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
    // ??? style:auto-update 19.467,
    // ??? style:class 19.470,
    // ignore style:data-style-name 19.473,
    // ??? style:default-outlinelevel 19.474,
    // ignore style:display-name 19.476,
    // ok style:family 19.480,
    // ignore style:list-level 19.499,
    // ignore style:list-style-name 19.500,
    // ignore style:master-page-name 19.501,
    // ok style:name 19.502,
    // ignore style:next-style-name 19.503,
    // ignore style:parent-style-name 19.510,
    // ignore style:percentage-data-style-name 19.511.
    attr: AttrMap2,
    /// Table style properties
    // ignore these attributes for now.
    graphicstyle: AttrMap2,
}

impl GraphicStyle {
    // Empty.
    pub(crate) fn new_empty() -> Self {
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

    /// Reference to this style.
    pub fn style_ref(&self) -> GraphicStyleRef {
        GraphicStyleRef::from(self.name())
    }

    /// Origin of the style.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Origin of the style.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Usage of the style.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Usage of the style.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Stylename.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Stylename.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// General attributes.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Graphic style attributes.
    pub(crate) fn graphicstyle(&self) -> &AttrMap2 {
        &self.graphicstyle
    }

    /// Graphic style attributes.
    pub(crate) fn graphicstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.graphicstyle
    }
}
