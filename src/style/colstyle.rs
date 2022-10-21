use std::fmt::{Display, Formatter};

use crate::attrmap2::AttrMap2;
use crate::style::units::{Length, PageBreak};
use crate::style::ParseStyleAttr;
use crate::style::Style;
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
    // ok fo:break-after 20.184,
    // ok fo:break-before 20.185,
    // ok style:column-width 20.254,
    // ok style:rel-column-width 20.338
    // ok style:use-optimal-column-width 20.393
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

    fo_break!(colstyle_mut);

    /// The style:rel-column-width attribute specifies a relative width of a column with a number
    /// value, followed by a ”*” (U+002A, ASTERISK) character. If rc is the relative with of the column, rs
    /// the sum of all relative columns widths, and ws the absolute width that is available for these
    /// columns the absolute width wc of the column is wc=rcws/rs.
    pub fn set_rel_col_width(&mut self, rel: f64) {
        self.colstyle
            .set_attr("style:rel-column-width", rel_width_string(rel));
    }

    /// The style:column-width attribute specifies a fixed width for a column.
    pub fn set_col_width(&mut self, width: Length) {
        if width == Length::Default {
            self.colstyle.clear_attr("style:column-width");
        } else {
            self.colstyle
                .set_attr("style:column-width", width.to_string());
        }
    }

    /// Parses the column width.
    pub fn col_width(&self) -> Result<Length, OdsError> {
        Length::parse_attr_def(self.colstyle.attr("style:column-width"), Length::Default)
    }

    /// The style:use-optimal-column-width attribute specifies that a column width should be
    /// recalculated automatically if content in the column changes.
    pub fn set_use_optimal_col_width(&mut self, opt: bool) {
        self.colstyle
            .set_attr("style:use-optimal-column-width", opt.to_string());
    }

    /// Parses the flag.
    pub fn use_optimal_col_width(&self) -> Result<bool, OdsError> {
        bool::parse_attr_def(self.colstyle.attr("style:use-optimal-column-width"), false)
    }
}
