use crate::attrmap2::AttrMap2;
use crate::style::{color_string, PageBreak, StyleOrigin, StyleUse, TextKeep};
use crate::Length;
use color::Rgb;

#[derive(Debug, Clone)]
pub struct TableRowStyle {
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
    tablerow_style: AttrMap2,
}

impl TableRowStyle {
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            tablerow_style: Default::default(),
        }
    }

    pub fn new<S: Into<String>>(name: S) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            tablerow_style: Default::default(),
        };
        s.set_name(name.into());
        s
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

    pub fn name(&self) -> Option<&String> {
        self.attr.attr("style:name")
    }

    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:name", name.into());
    }

    pub fn attr(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attr_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    pub fn tablerow_style(&self) -> &AttrMap2 {
        &self.tablerow_style
    }

    pub fn tablerow_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.tablerow_style
    }

    fo_background_color!(tablerow_style_mut);
    fo_break!(tablerow_style_mut);
    fo_keep_together!(tablerow_style_mut);

    pub fn set_min_row_height(&mut self, min_height: Length) {
        self.tablerow_style
            .set_attr("style:min-row-height", min_height.to_string());
    }

    pub fn set_row_height(&mut self, height: Length) {
        self.tablerow_style
            .set_attr("style:row-height", height.to_string());
    }

    pub fn set_use_optimal_row_height(&mut self, opt: bool) {
        self.tablerow_style
            .set_attr("style:use-optimal-row-height", opt.to_string());
    }
}
