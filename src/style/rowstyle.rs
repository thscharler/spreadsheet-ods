use crate::attrmap2::AttrMap2;
use crate::style::units::{Length, PageBreak, TextKeep};
use crate::style::{color_string, StyleOrigin, StyleUse};
use color::Rgb;

style_ref!(RowStyleRef);

/// Describes the style information for a table row.
/// Hardly ever used. It's easier to set the row_height via
/// Sheet::set_row_height.
///
#[derive(Debug, Clone)]
pub struct RowStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
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
    rowstyle: AttrMap2,
}

impl RowStyle {
    /// empty
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            rowstyle: Default::default(),
        }
    }

    /// New Style.
    pub fn new<S: Into<String>>(name: S) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            rowstyle: Default::default(),
        };
        s.set_name(name.into());
        s
    }

    /// Reference to the style.
    pub fn style_ref(&self) -> RowStyleRef {
        RowStyleRef::from(self.name().unwrap().clone())
    }

    /// Origin. Should always be Content.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Origin. Should always be Content.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Usage. Should always be Automatic.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Usage. Should always be Automatic.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Name
    pub fn name(&self) -> Option<&String> {
        self.attr.attr("style:name")
    }

    /// Name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:name", name.into());
    }

    /// General attributes.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Style attributes.
    pub fn rowstyle(&self) -> &AttrMap2 {
        &self.rowstyle
    }

    /// Style attributes.
    pub fn rowstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.rowstyle
    }

    fo_background_color!(rowstyle_mut);
    fo_break!(rowstyle_mut);
    fo_keep_together!(rowstyle_mut);

    /// Minimum row-height.
    pub fn set_min_row_height(&mut self, min_height: Length) {
        self.rowstyle
            .set_attr("style:min-row-height", min_height.to_string());
    }

    /// Fixed row-height.
    pub fn set_row_height(&mut self, height: Length) {
        self.rowstyle
            .set_attr("style:row-height", height.to_string());
    }

    /// Optimal row-height.
    pub fn set_use_optimal_row_height(&mut self, opt: bool) {
        self.rowstyle
            .set_attr("style:use-optimal-row-height", opt.to_string());
    }
}
