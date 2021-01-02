use crate::attrmap2::AttrMap2;
use crate::style::{
    color_string, percent_string, shadow_string, FontStyle, FontWeight, Length, LineMode,
    LineStyle, LineType, LineWidth, StyleOrigin, StyleUse, TextPosition, TextRelief, TextTransform,
};
use color::Rgb;

#[derive(Debug, Clone)]
pub struct TextStyle {
    /// From where did we get this style.
    origin: StyleOrigin,
    /// Which tag contains this style.
    styleuse: StyleUse,
    /// General attributes
    // ??? style:auto-update 19.467,
    // ??? style:class 19.470,
    // ignore style:data-style-name 19.473,
    // ignore style:default-outlinelevel 19.474,
    // ok style:display-name 19.476,
    // ok style:family 19.480,
    // ignore style:list-level 19.499,
    // ignore style:list-style-name 19.500,
    // ignore style:master-page-name 19.501,
    // ok style:name 19.502,
    // ignore style:next-style-name 19.503,
    // ok style:parent-style-name 19.510,
    // ignore style:percentage-data-style-name 19.511.
    attr: AttrMap2,
    text_style: AttrMap2,
}

impl TextStyle {
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            text_style: Default::default(),
        }
    }

    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            text_style: Default::default(),
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

    pub fn display_name(&self) -> Option<&String> {
        self.attr.attr("style:display-name")
    }

    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:display-name", name.into());
    }

    pub fn parent_style(&self) -> Option<&String> {
        self.attr.attr("style:parent-style-name")
    }

    pub fn set_parent_style<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:parent-style-name", name.into());
    }

    pub fn attr(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attr_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    pub fn text_style(&self) -> &AttrMap2 {
        &self.text_style
    }

    pub fn text_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.text_style
    }

    text!(text_style_mut);
}
