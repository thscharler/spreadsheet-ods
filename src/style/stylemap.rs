//!
//! Conditional styles.
//!

use crate::condition::{Condition, ValueCondition};
use crate::CellRef;

/// A style-map is one way for conditional formatting of cells.
///
/// It seems this is always translated into calcext:conditional-formats
/// which seem to be the preferred way to deal with this. But it still
/// works somewhat.
#[derive(Clone, Debug, Default)]
pub struct StyleMap {
    condition: String,
    applied_style: String,
    base_cell: Option<CellRef>,
}

impl StyleMap {
    ///  Create a stylemap. When the condition is fullfilled the style
    /// applied_style is used. The base_cell is used to resolve all relative
    /// cell-references within the condition.
    pub fn new<T: Into<String>>(
        condition: Condition,
        applied_style: T,
        base_cell: Option<CellRef>,
    ) -> Self {
        Self {
            condition: condition.to_string(),
            applied_style: applied_style.into(),
            base_cell: base_cell,
        }
    }

    /// Condition
    pub fn condition(&self) -> &String {
        &self.condition
    }

    /// Condition
    pub fn set_condition(&mut self, cond: ValueCondition) {
        self.condition = cond.to_string();
    }

    /// The applied style.
    pub fn applied_style(&self) -> &String {
        &self.applied_style
    }

    /// Sets the applied style.
    pub fn set_applied_style<S: Into<String>>(&mut self, style: S) {
        self.applied_style = style.into();
    }

    /// Base cell.
    pub fn base_cell(&self) -> Option<&CellRef> {
        self.base_cell.as_ref()
    }

    /// Sets the base cell.
    pub fn set_base_cell(&mut self, cellref: Option<CellRef>) {
        self.base_cell = cellref;
    }
}
