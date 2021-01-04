use crate::attrmap2::AttrMap2;
use crate::style::units::{Length, PageBreak};
use crate::style::{rel_width_string, StyleOrigin, StyleUse};

style_ref!(ColStyleRef);

/// Describes the style information for a table column.
/// Hardly ever used. It's easier to set the col_width via
/// Sheet::set_col_width
///
#[derive(Debug, Clone)]
pub struct ColStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// General attributes
    // ??? style:auto-update 19.467,
    // ??? style:class 19.470,
    // ignore style:data-style-name 19.473,
    // ignore style:default-outlinelevel 19.474,
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
    colstyle: AttrMap2,
}

impl ColStyle {
    /// empty
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            colstyle: Default::default(),
        }
    }

    /// New Style.
    pub fn new<S: Into<String>>(name: S) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            colstyle: Default::default(),
        };
        s.set_name(name.into());
        s
    }

    /// Returns a reference.
    pub fn style_ref(&self) -> ColStyleRef {
        ColStyleRef::from(self.name().unwrap().clone())
    }

    /// Origin. Should always be Content.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Origin. Should always be Content.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Should always be Automatic.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Should always be Automatic.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Name.
    pub fn name(&self) -> Option<&String> {
        self.attr.attr("style:name")
    }

    /// Name.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:name", name.into());
    }

    /// Attributes
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Attributes
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Style attributes
    pub fn colstyle(&self) -> &AttrMap2 {
        &self.colstyle
    }

    /// Style attributes
    pub fn colstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.colstyle
    }

    fo_break!(colstyle_mut);

    /// Relative weights for the column width
    pub fn set_rel_col_width(&mut self, rel: f64) {
        self.colstyle
            .set_attr("style:rel-column-width", rel_width_string(rel));
    }

    /// Column width
    pub fn set_col_width(&mut self, width: Length) {
        self.colstyle
            .set_attr("style:column-width", width.to_string());
    }

    /// Override switch for the column width.
    pub fn set_use_optimal_col_width(&mut self, opt: bool) {
        self.colstyle
            .set_attr("style:use-optimal-column-width", opt.to_string());
    }
}
