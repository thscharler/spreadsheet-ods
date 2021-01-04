//! Styles define a large number of attributes. These are grouped together
//! as table, row, column, cell, paragraph and text attributes.
//!
//! ```
//! use spreadsheet_ods::{CellRef, WorkBook};
//! use spreadsheet_ods::style::{StyleOrigin, StyleUse, CellStyle};
//! use color::Rgb;
//! use spreadsheet_ods::style::stylemap::StyleMap;
//!
//! let mut wb = WorkBook::new();
//!
//! let mut st = CellStyle::new("ce12", &"num2".into());
//! st.set_color(Rgb::new(192, 128, 0));
//! st.set_font_bold();
//! wb.add_cellstyle(st);
//!
//! let mut st = CellStyle::new("ce11", &"num2".into());
//! st.set_color(Rgb::new(0, 192, 128));
//! st.set_font_bold();
//! wb.add_cellstyle(st);
//!
//! let mut st = CellStyle::new("ce13", &"num4".into());
//! st.push_stylemap(StyleMap::new("cell-content()=\"BB\"", "ce12", CellRef::remote("sheet0", 4, 3)));
//! st.push_stylemap(StyleMap::new("cell-content()=\"CC\"", "ce11", CellRef::remote("sheet0", 4, 3)));
//! wb.add_cellstyle(st);
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

mod cellstyle;
mod colstyle;
mod fontface;
mod graphicstyle;
mod masterpage;
mod pagestyle;
mod paragraphstyle;
mod rowstyle;
pub mod stylemap;
mod tablestyle;
pub mod tabstop;
mod textstyle;
pub mod units;

pub use crate::attrmap2::*;
pub use cellstyle::*;
pub use colstyle::*;
pub use fontface::*;
pub use graphicstyle::*;
pub use masterpage::*;
pub use pagestyle::*;
pub use paragraphstyle::*;
pub use rowstyle::*;
pub use tablestyle::*;
pub use textstyle::*;

use crate::style::units::{Border, Length};
use color::Rgb;

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
