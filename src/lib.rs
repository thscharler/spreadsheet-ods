#![doc = include_str!("../crate.md")]

#[macro_use]
mod macro_attr_draw;
#[macro_use]
mod macro_attr_style;
#[macro_use]
mod macro_attr_fo;
#[macro_use]
mod macro_attr_svg;
#[macro_use]
mod macro_attr_text;
#[macro_use]
mod macro_attr_number;
#[macro_use]
mod macro_attr_table;
#[macro_use]
mod macro_attr_xlink;
#[macro_use]
mod unit_macro;
#[macro_use]
mod format_macro;
#[macro_use]
mod style_macro;
#[macro_use]
mod text_macro;

mod attrmap2;
mod cell;
mod config;
mod ds;
mod io;
mod locale;
mod sheet;
#[macro_use]
mod value;
mod workbook;

pub mod condition;
pub mod defaultstyles;
pub mod draw;
pub mod error;
pub mod format;
#[macro_use]
pub mod formula;
pub mod manifest;
pub mod metadata;
pub mod refs;
pub mod style;
pub mod text;
pub mod validation;
pub mod xlink;
pub mod xmltree;

pub use crate::error::OdsError;
pub use crate::format::{
    ValueFormatBoolean, ValueFormatCurrency, ValueFormatDateTime, ValueFormatNumber,
    ValueFormatPercentage, ValueFormatRef, ValueFormatText, ValueFormatTimeDuration,
};
pub use crate::io::read::{
    read_fods, read_fods_buf, read_fods_from, read_ods, read_ods_buf, read_ods_from, OdsOptions,
};
pub use crate::io::write::{
    write_fods, write_fods_buf, write_fods_to, write_ods, write_ods_buf,
    write_ods_buf_uncompressed, write_ods_to,
};
pub use crate::refs::{CellRange, CellRef, ColRange, RowRange};
pub use crate::style::units::{Angle, Length};
pub use crate::style::{CellStyle, CellStyleRef};
pub use cell::{CellContent, CellContentRef, CellSpan};
pub use color;
pub use sheet::{CellIter, Grouped, Range, Sheet, SheetConfig, SplitMode, Visibility};
pub use value::{Value, ValueType};
pub use workbook::{EventListener, Script, WorkBook, WorkBookConfig};

// Use the IndexMap for debugging, makes diffing much easier.
// Otherwise the std::HashMap is good.
// pub(crate) type HashMap<K, V> = indexmap::IndexMap<K, V>;
// pub(crate) type HashMapIter<'a, K, V> = indexmap::map::Iter<'a, K, V>;
pub(crate) type HashMap<K, V> = std::collections::HashMap<K, V>;
pub(crate) type HashMapIter<'a, K, V> = std::collections::hash_map::Iter<'a, K, V>;
