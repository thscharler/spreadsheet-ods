use crate::attrmap2::AttrMap2;
use crate::style::units::{
    FontStyle, FontWeight, Length, LineMode, LineStyle, LineType, LineWidth, TextPosition,
    TextRelief, TextTransform,
};
use crate::style::{color_string, percent_string, shadow_string, StyleOrigin, StyleUse};
use color::Rgb;

style_ref!(TextStyleRef);

/// Text style.
/// This is not used for cell-formatting. Use CellStyle instead.
///
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
    textstyle: AttrMap2,
}

impl TextStyle {
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            textstyle: Default::default(),
        }
    }

    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            textstyle: Default::default(),
        };
        s.set_name(name.into());
        s
    }

    pub fn style_ref(&self) -> TextStyleRef {
        TextStyleRef::from(self.name().unwrap().clone())
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

    pub fn set_parent_style(&mut self, name: &TextStyleRef) {
        self.attr
            .set_attr("style:parent-style-name", name.to_string());
    }

    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    pub fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    pub fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    text!(textstyle_mut);
}
