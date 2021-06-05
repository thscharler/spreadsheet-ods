use crate::condition::ValueCondition;
use crate::CellRef;

/// One style mapping.
///
/// The rules for this are not very clear. It writes the necessary data fine,
/// but the interpretation by LibreOffice is not very intelligible.
///
/// * The base-cell must include a table-name.
/// * LibreOffice always adds calcext:conditional-formats which I can't handle.
///
/// TODO: clarify all of this.
///
#[derive(Clone, Debug, Default)]
pub struct StyleMap {
    condition: String,
    applied_style: String,
    base_cell: CellRef,
}

impl StyleMap {
    pub fn new<T: Into<String>>(
        condition: ValueCondition,
        applied_style: T,
        base_cell: CellRef,
    ) -> Self {
        Self {
            condition: condition.to_string(),
            applied_style: applied_style.into(),
            base_cell,
        }
    }

    pub fn condition(&self) -> &String {
        &self.condition
    }

    pub fn set_condition(&mut self, cond: ValueCondition) {
        self.condition = cond.to_string();
    }

    pub fn applied_style(&self) -> &String {
        &self.applied_style
    }

    pub fn set_applied_style<S: Into<String>>(&mut self, style: S) {
        self.applied_style = style.into();
    }

    pub fn base_cell(&self) -> &CellRef {
        &self.base_cell
    }

    pub fn set_base_cell(&mut self, cellref: CellRef) {
        self.base_cell = cellref;
    }
}
