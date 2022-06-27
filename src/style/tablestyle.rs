use crate::attrmap2::AttrMap2;
use crate::style::units::{Length, PageBreak, TextKeep, WritingMode};
use crate::style::{color_string, shadow_string, MasterPageRef, StyleOrigin, StyleUse};
use color::Rgb;
use std::fmt::{Display, Formatter};

style_ref!(TableStyleRef);

/// Describes the style information for a table.
///
#[derive(Debug, Clone)]
pub struct TableStyle {
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
    // ok style:master-page-name 19.501,
    // ok style:name 19.502,
    // ignore style:next-style-name 19.503,
    // ignore style:parent-style-name 19.510,
    // ignore style:percentage-data-style-name 19.511.
    attr: AttrMap2,
    /// Table style properties
    tablestyle: AttrMap2,
}

impl TableStyle {
    /// empty
    pub(crate) fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            tablestyle: Default::default(),
        }
    }

    /// Creates a new Style.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            tablestyle: Default::default(),
        }
    }

    /// Style reference.
    pub fn style_ref(&self) -> TableStyleRef {
        TableStyleRef::from(self.name())
    }

    /// Origin of the style.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Origin of the style.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Designation of the style.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Designation of the style.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Style name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Style name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Sets the reference to the pageformat.
    pub fn set_master_page_name(&mut self, name: &MasterPageRef) {
        self.attr
            .set_attr("style:master-page-name", name.to_string());
    }

    /// Reference to the pageformat.
    pub fn master_page_name(&self) -> Option<&String> {
        self.attr.attr("style:master-page-name")
    }

    /// Access to all stored attributes.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Access to all stored attributes.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Access to all style attributes.
    pub fn tablestyle(&self) -> &AttrMap2 {
        &self.tablestyle
    }

    /// Access to all style attributes.
    pub fn tablestyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.tablestyle
    }

    // style:may-break-between-rows 20.319,
    // style:page-number 20.328,
    // style:rel-width 20.340,
    // style:width 20.399,
    // table:align 20.414,
    // table:border-model 20.415,
    // table:display 20.416
    // table:tab-color 19.731.

    fo_background_color!(tablestyle_mut);
    fo_break!(tablestyle_mut);
    fo_keep_with_next!(tablestyle_mut);
    fo_margin!(tablestyle_mut);
    style_shadow!(tablestyle_mut);
    style_writing_mode!(tablestyle_mut);
}
