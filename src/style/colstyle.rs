use std::fmt::{Display, Formatter};

use crate::attrmap2::AttrMap2;
use crate::style::units::{Length, PageBreak};
use crate::style::ParseStyleAttr;
use crate::style::{rel_width_string, StyleOrigin, StyleUse};
use crate::OdsError;

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
    // ignore style:master-page-name 19.501, => PARAGRAPH, TABLE
    // ignore style:name 19.502, Not mapped as an attribute.
    // ignore style:next-style-name 19.503, => PARAGRAPH
    // ok style:parent-style-name 19.510 => ALL
    // na style:percentage-data-style-name 19.511. => PARAGRAPH?
    attr: AttrMap2,
    /// Column style properties
    colstyle: AttrMap2,
}

styles_styles!(ColStyle, ColStyleRef);

impl ColStyle {
    /// empty
    pub(crate) fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            colstyle: Default::default(),
        }
    }

    /// New Style.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            colstyle: Default::default(),
        }
    }

    /// Attributes
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Attributes
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Style attributes
    pub(crate) fn colstyle(&self) -> &AttrMap2 {
        &self.colstyle
    }

    /// Style attributes
    pub(crate) fn colstyle_mut(&mut self) -> &mut AttrMap2 {
        &mut self.colstyle
    }

    fo_break!(colstyle);
    style_column_width!(colstyle);
    style_rel_column_width!(colstyle);
    style_use_optimal_column_width!(colstyle);
}
