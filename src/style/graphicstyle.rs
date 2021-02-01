use crate::attrmap2::AttrMap2;
use crate::style::{StyleOrigin, StyleUse};

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
    graphicstyle: AttrMap2,
}

impl GraphicStyle {
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            graphicstyle: Default::default(),
        }
    }

    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            graphicstyle: Default::default(),
        }
    }

    pub fn style_ref(&self) -> GraphicStyleRef {
        GraphicStyleRef::from(self.name())
    }

    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    pub fn graphicstyle(&self) -> &AttrMap2 {
        &self.graphicstyle
    }

    pub fn graphicstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.graphicstyle
    }
}
