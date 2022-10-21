use color::Rgb;

use crate::attrmap2::AttrMap2;
use crate::style::units::{Length, PageBreak, TextKeep};
use crate::style::ParseStyleAttr;
use crate::style::Style;
use crate::style::{color_string, StyleOrigin, StyleUse};
use crate::OdsError;
use std::fmt::{Display, Formatter};

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
    // ignore style:name 19.502, => Not mapped as an attribute.
    // ignore style:next-style-name 19.503, => PARAGRAPH
    // ok style:parent-style-name 19.510 => ALL
    // ignore style:percentage-data-style-name 19.511. => PARAGRAPH?
    attr: AttrMap2,
    /// Table style properties
    rowstyle: AttrMap2,
}

styles_styles!(RowStyle, RowStyleRef);

impl RowStyle {
    /// empty
    pub(crate) fn new_empty() -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: Default::default(),
            attr: Default::default(),
            rowstyle: Default::default(),
        }
    }

    /// New Style.
    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            origin: Default::default(),
            styleuse: Default::default(),
            name: name.into(),
            attr: Default::default(),
            rowstyle: Default::default(),
        }
    }

    /// General attributes.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Style attributes.
    pub(crate) fn rowstyle(&self) -> &AttrMap2 {
        &self.rowstyle
    }

    /// Style attributes.
    pub(crate) fn rowstyle_mut(&mut self) -> &mut AttrMap2 {
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

    /// Parses the row height
    pub fn row_height(&self) -> Result<Length, OdsError> {
        Length::parse_attr_def(self.rowstyle.attr("style:row-height"), Length::Default)
    }

    /// Optimal row-height.
    pub fn set_use_optimal_row_height(&mut self, opt: bool) {
        self.rowstyle
            .set_attr("style:use-optimal-row-height", opt.to_string());
    }

    /// Parses the flag.
    pub fn use_optimal_row_height(&self) -> Result<bool, OdsError> {
        bool::parse_attr_def(self.rowstyle.attr("style:use-optimal-row-height"), false)
    }
}
