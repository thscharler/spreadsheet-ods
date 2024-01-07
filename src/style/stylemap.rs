//!
//! Conditional styles.
//!

use crate::condition::Condition;
use crate::CellRef;
use get_size::GetSize;
use get_size_derive::GetSize;

/// A style-map is one way for conditional formatting of cells.
///
/// It seems this is always translated into calcext:conditional-formats
/// which seem to be the preferred way to deal with this. But it still
/// works somewhat.
#[derive(Clone, Debug, Default, GetSize)]
pub struct StyleMap {
    condition: Condition,
    applied_style: String, // todo: general style
    base_cell: Option<CellRef>,
}

impl StyleMap {
    ///  Create a stylemap. When the condition is fullfilled the style
    /// applied_style is used. The base_cell is used to resolve all relative
    /// cell-references within the condition.
    pub fn new<T: AsRef<str>>(
        condition: Condition,
        applied_style: T,
        base_cell: Option<CellRef>,
    ) -> Self {
        Self {
            condition,
            applied_style: applied_style.as_ref().to_string(),
            base_cell,
        }
    }

    /// Condition
    pub fn condition(&self) -> &Condition {
        &self.condition
    }

    /// Condition
    pub fn set_condition(&mut self, cond: Condition) {
        self.condition = cond;
    }

    /// The applied style.
    pub fn applied_style(&self) -> &String {
        &self.applied_style
    }

    /// Sets the applied style.
    pub fn set_applied_style<S: AsRef<str>>(&mut self, style: S) {
        self.applied_style = style.as_ref().to_string();
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
