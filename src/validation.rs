use std::convert::TryFrom;

use crate::condition::Condition;
use crate::{CellRef, OdsError};

/// This defines how lists of entries are displayed to the user.
#[derive(Copy, Clone, Debug)]
pub enum ValidationDisplay {
    NoDisplay,
    Unsorted,
    SortAscending,
}

impl Default for ValidationDisplay {
    fn default() -> Self {
        ValidationDisplay::Unsorted
    }
}

impl TryFrom<&str> for ValidationDisplay {
    type Error = OdsError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "unsorted" => Ok(ValidationDisplay::Unsorted),
            "sort-ascending" => Ok(ValidationDisplay::SortAscending),
            "none" => Ok(ValidationDisplay::NoDisplay),
            _ => Err(OdsError::Parse(
                format!("unknown value or table:display-list: {}", value).to_string(),
            )),
        }
    }
}

style_ref!(ValidationRef);

/// Cell content validations.
/// This defines a validity constraint via the contained condition.
/// It can be applied to a cell by setting the validation name.
#[derive(Clone, Debug, Default)]
pub struct Validation {
    name: String,
    condition: String,
    base_cell: CellRef,
    allow_empty: bool,
    display_list: ValidationDisplay,
}

impl Validation {
    pub fn new() -> Self {
        Self {
            name: "".to_string(),
            condition: "".to_string(),
            base_cell: CellRef::new(),
            allow_empty: true,
            display_list: Default::default(),
        }
    }

    /// Validation name.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Validation name.
    pub fn name(&self) -> &str {
        self.name.as_str()
    }

    /// Creates a reference struct for this one.
    pub fn validation_ref(&self) -> ValidationRef {
        ValidationRef::from(self.name.clone())
    }

    /// Sets the condition that is checked for new values.
    pub fn set_condition(&mut self, cond: Condition) {
        self.condition = cond.to_string();
    }

    /// Condition for new values.
    pub fn condition(&self) -> &str {
        self.condition.as_str()
    }

    /// Base-cell for the validation. Relative CellReferences in the
    /// condition are relative to this cell. They are moved with the
    /// actual cell this condition is applied to.
    pub fn set_base_cell(&mut self, base: CellRef) {
        self.base_cell = base.into();
    }

    /// Base-cell for the validation.
    pub fn base_cell(&self) -> &CellRef {
        &self.base_cell
    }

    /// Empty ok?
    pub fn set_allow_empty(&mut self, allow: bool) {
        self.allow_empty = allow;
    }

    /// Empty ok?
    pub fn allow_empty(&self) -> bool {
        self.allow_empty
    }

    /// Display list of choices.
    pub fn set_display(&mut self, display: ValidationDisplay) {
        self.display_list = display;
    }

    /// Display list of choices.
    pub fn display(&self) -> ValidationDisplay {
        self.display_list
    }
}
