use crate::attrmap2::AttrMap2;
use crate::style::units::{
    Length, Margin, PageBreak, PageNumber, RelativeWidth, TableAlign, TableBorderModel, TextKeep,
    WritingMode,
};
use crate::style::Style;
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
    // ok style:auto-update 19.467 => ALL
    // ok style:class 19.470, => ALL
    // ignore style:data-style-name 19.473, => CELL, CHART
    // ignore style:default-outlinelevel 19.474, => PARAGRAPH
    // ok style:display-name 19.476, => ALL
    // ignore style:family 19.480, => Not mapped as an attribute.
    // ignore style:list-level 19.499, => PARAGRAPH
    // ignore style:list-style-name 19.500, => PARAGRAPH
    // ok style:master-page-name 19.501, => PARAGRAPH, TABLE
    // ignore style:name 19.502, => Not mapped as an attribute.
    // ignore style:next-style-name 19.503, => PARAGRAPH
    // ok style:parent-style-name 19.510 => ALL
    // ignore style:percentage-data-style-name 19.511. => PARAGRAPH?
    attr: AttrMap2,
    /// Table style properties
    tablestyle: AttrMap2,
}

styles_styles!(TableStyle, TableStyleRef);

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

    style_master_page!(attrmap_mut);

    /// Access to all stored attributes.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Access to all stored attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Access to all style attributes.
    pub(crate) fn tablestyle(&self) -> &AttrMap2 {
        &self.tablestyle
    }

    /// Access to all style attributes.
    pub(crate) fn tablestyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.tablestyle
    }

    fo_background_color!(tablestyle_mut);
    fo_break!(tablestyle_mut);
    fo_keep_with_next!(tablestyle_mut);
    fo_margin!(tablestyle_mut);
    style_may_break_between_rows!(tablestyle);
    style_page_number!(tablestyle_mut);
    style_rel_width!(tablestyle);
    style_width!(tablestyle);
    style_shadow!(tablestyle_mut);
    style_writing_mode!(tablestyle_mut);

    table_align!(tablestyle);
    table_border_model!(tablestyle);
    table_display!(tablestyle);
    table_tab_color!(tablestyle);
}
