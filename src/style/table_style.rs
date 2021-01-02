use crate::attrmap2::AttrMap2;
use crate::style::{
    color_string, shadow_string, AnyStyle, Length, PageBreak, StyleOrigin, StyleUse, TextKeep,
    WritingMode,
};
use color::Rgb;

#[derive(Debug, Clone)]
pub struct TableStyle {
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
    // ignorestyle:list-style-name 19.500,
    // ok style:master-page-name 19.501,
    // ok style:name 19.502,
    // ignore style:next-style-name 19.503,
    // ignore style:parent-style-name 19.510,
    // ignore style:percentage-data-style-name 19.511.
    attr: AttrMap2,
    /// Table style properties
    table_style: AttrMap2,
}

impl TableStyle {
    pub fn empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            table_style: Default::default(),
        }
    }

    pub fn new<S: Into<String>, T: Into<String>>(name: S) -> Self {
        let mut s = Self {
            origin: Default::default(),
            styleuse: Default::default(),
            attr: Default::default(),
            table_style: Default::default(),
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

    /// Sets the value format.
    pub fn set_master_page_name<S: Into<String>>(&mut self, name: S) {
        self.attr.set_attr("style:master-page-name", name.into());
    }

    /// Returns the value format.
    pub fn master_page_name(&self) -> Option<&String> {
        self.attr.attr("style:master-page-name")
    }

    pub fn attr(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attr_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    pub fn table_style(&self) -> &AttrMap2 {
        &self.table_style
    }

    pub fn table_style_mut(&mut self) -> &mut AttrMap2 {
        &mut self.table_style
    }

    fo_background_color!(table_style_mut);
    fo_margin!(table_style_mut);
    fo_break!(table_style_mut);
    fo_keep_with_next!(table_style_mut);
    style_shadow!(table_style_mut);
    style_writing_mode!(table_style_mut);
}

impl From<TableStyle> for AnyStyle {
    fn from(style: TableStyle) -> Self {
        AnyStyle::TableStyle(style)
    }
}
