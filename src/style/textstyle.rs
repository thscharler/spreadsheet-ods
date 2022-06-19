use crate::attrmap2::AttrMap2;
use crate::style::units::{
    FontStyle, FontWeight, Length, LineMode, LineStyle, LineType, LineWidth, TextPosition,
    TextRelief, TextTransform,
};
use crate::style::{color_string, percent_string, shadow_string, StyleOrigin, StyleUse};
use color::Rgb;
use std::fmt::{Display, Formatter};

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
    /// Style name
    name: String,
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
    /// Empty.
    pub(crate) fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            textstyle: Default::default(),
        }
    }

    /// A new named style.
    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            textstyle: Default::default(),
        }
    }

    /// Creates a reference to this style.
    pub fn style_ref(&self) -> TextStyleRef {
        TextStyleRef::from(self.name())
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

    /// Stylename
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Stylename
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Display name for named styles.
    pub fn display_name(&self) -> Option<&String> {
        self.attr.attr("style:display-name")
    }

    /// Display name for named styles.
    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:display-name", name.into());
    }

    /// Parent style.
    pub fn parent_style(&self) -> Option<&String> {
        self.attr.attr("style:parent-style-name")
    }

    /// Parent style.
    pub fn set_parent_style(&mut self, name: &TextStyleRef) {
        self.attr
            .set_attr("style:parent-style-name", name.to_string());
    }

    /// General attributes for the style.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes for the style.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// All text-attributes for the style.
    pub fn textstyle(&self) -> &AttrMap2 {
        &self.textstyle
    }

    /// All text-attributes for the style.
    pub fn textstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.textstyle
    }

    text!(textstyle_mut);
}
