//! Styles define a large number of attributes. These are grouped together
//! as table, row, column, cell, paragraph and text attributes.
//!
//! ```
//! use spreadsheet_ods::{CellRef, WorkBook};
//! use spreadsheet_ods::style::{StyleOrigin, StyleUse, AttrText, StyleMap, CellStyle};
//! use color::Rgb;
//!
//! let mut wb = WorkBook::new();
//!
//! let mut st = CellStyle::new("ce12", "num2");
//! st.set_color(Rgb::new(192, 128, 0));
//! st.set_font_bold();
//! wb.add_cell_style(st);
//!
//! let mut st = CellStyle::new("ce11", "num2");
//! st.set_color(Rgb::new(0, 192, 128));
//! st.set_font_bold();
//! wb.add_cell_style(st);
//!
//! let mut st = CellStyle::new("ce13", "num4");
//! st.push_stylemap(StyleMap::new("cell-content()=\"BB\"", "ce12", CellRef::remote("sheet0", 4, 3)));
//! st.push_stylemap(StyleMap::new("cell-content()=\"CC\"", "ce11", CellRef::remote("sheet0", 4, 3)));
//! wb.add_cell_style(st);
//! ```
//! Styles can be defined in content.xml or as global styles in styles.xml. This
//! is reflected as the StyleOrigin. The StyleUse differentiates between automatic
//! and user visible, named styles. And third StyleFor defines for which part of
//! the document the style can be used.
//!
//! Cell styles usually reference a value format for text formatting purposes.
//!
//! Styles can also link to a parent style and to a pagelayout.
//!

mod attr;
#[macro_use]
mod attr_macro;
mod cell_style;
mod column_style;
mod fontface;
mod graphic_style;
mod pagelayout;
mod paragraph_style;
mod row_style;
mod stylemap;
mod table_style;
mod tabstop;
mod text_style;
mod units;

pub use crate::attrmap::*;
pub use attr::*;
pub use cell_style::*;
pub use column_style::*;
pub use fontface::*;
pub use graphic_style::*;
pub use pagelayout::*;
pub use paragraph_style::*;
pub use row_style::*;
pub use stylemap::*;
pub use table_style::*;
pub use tabstop::*;
pub use text_style::*;
pub use units::*;

use crate::sealed::Sealed;
use color::Rgb;
use string_cache::DefaultAtom;

/// Origin of a style. Content.xml or Styles.xml.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleOrigin {
    Content,
    Styles,
}

impl Default for StyleOrigin {
    fn default() -> Self {
        StyleOrigin::Content
    }
}

/// Placement of a style. office:styles or office:automatic-styles
/// Defines the usage pattern for the style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleUse {
    Default,
    Named,
    Automatic,
}

impl Default for StyleUse {
    fn default() -> Self {
        StyleUse::Automatic
    }
}

/// Applicability of this style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleFor {
    Table,
    TableRow,
    TableColumn,
    TableCell,
    Graphic,
    Paragraph,
    Text,
    None,
}

impl Default for StyleFor {
    fn default() -> Self {
        StyleFor::None
    }
}

/// Text styles.
#[derive(Clone, Debug, Default)]
pub struct TextAttr {
    attr: AttrMapType,
}

impl Sealed for TextAttr {}

impl AttrMap for TextAttr {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl<'a> IntoIterator for &'a TextAttr {
    type Item = (&'a DefaultAtom, &'a String);
    type IntoIter = AttrMapIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        AttrMapIter::from(self.attr_map())
    }
}

impl AttrFoBackgroundColor for TextAttr {}

impl AttrText for TextAttr {}

pub(crate) fn color_string(color: Rgb<u8>) -> String {
    format!("#{:02x}{:02x}{:02x}", color.r, color.g, color.b)
}

pub(crate) fn shadow_string(
    x_offset: Length,
    y_offset: Length,
    blur: Option<Length>,
    color: Rgb<u8>,
) -> String {
    if let Some(blur) = blur {
        format!("{} {} {} {}", color_string(color), x_offset, y_offset, blur)
    } else {
        format!("{} {} {}", color_string(color), x_offset, y_offset)
    }
}

pub(crate) fn rel_width_string(value: f64) -> String {
    format!("{}*", value)
}

pub(crate) fn border_string(width: Length, border: Border, color: Rgb<u8>) -> String {
    format!(
        "{} {} #{:02x}{:02x}{:02x}",
        width, border, color.r, color.g, color.b
    )
}

pub(crate) fn percent_string(value: f64) -> String {
    format!("{}%", value)
}

pub(crate) fn border_line_width_string(inner: Length, space: Length, outer: Length) -> String {
    format!("{} {} {}", inner, space, outer)
}
