//!
//! Conditional styles.
//!

use crate::condition::ValueCondition;

/// A style-map is one way for conditional formatting of value formats.
#[derive(Clone, Debug, Default)]
pub struct ValueStyleMap {
    condition: String,
    applied_style: String,
}

impl ValueStyleMap {
    /// Create a stylemap for a ValueFormat. When the condition is fullfilled the style
    /// applied_style is used.
    pub fn new<T: Into<String>>(condition: ValueCondition, applied_style: T) -> Self {
        Self {
            condition: condition.to_string(),
            applied_style: applied_style.into(),
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
}
