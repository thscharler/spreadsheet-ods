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
mod macro_units;
#[macro_use]
mod macro_format;
#[macro_use]
mod macro_style;
#[macro_use]
mod macro_text;

mod attrmap2;
mod cell_;
pub mod condition;
mod config;
pub mod defaultstyles;
pub mod draw;
mod ds;
mod error;
pub mod format;
#[macro_use]
pub mod formula;
mod io;
mod locale;
pub mod manifest;
pub mod metadata;
pub mod refs;
mod sheet_;
pub mod style;
pub mod text;
pub mod validation;
#[macro_use]
mod value_;
mod workbook_;
pub mod xlink;
pub mod xmltree;

pub use color;

pub use crate::cell_::{CellContent, CellContentRef};
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
/// Detail structs for a Cell.
pub mod cell {
    pub use crate::cell_::CellSpan;
}
pub use crate::sheet_::Sheet;
/// Detail structs for a Sheet.
pub mod sheet {
    pub use crate::sheet_::{CellIter, Grouped, Range, SheetConfig, SplitMode, Visibility};
}
pub use crate::value_::{Value, ValueType};
// pub mod value {
// }
pub use crate::workbook_::WorkBook;
/// Detail structs for the WorkBook.
pub mod workbook {
    pub use crate::workbook_::{EventListener, Script, WorkBookConfig};
}

// Use the IndexMap for debugging, makes diffing much easier.
// Otherwise the std::HashMap is good.
// pub(crate) type HashMap<K, V> = indexmap::IndexMap<K, V>;
// pub(crate) type HashMapIter<'a, K, V> = indexmap::map::Iter<'a, K, V>;
pub(crate) type HashMap<K, V> = std::collections::HashMap<K, V>;
pub(crate) type HashMapIter<'a, K, V> = std::collections::hash_map::Iter<'a, K, V>;
