use std::convert::TryFrom;
use std::fmt::{Display, Formatter};

use crate::condition::Condition;
use crate::text::TextTag;
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
            _ => Err(OdsError::Parse(format!(
                "unknown value or table:display-list: {}",
                value
            ))),
        }
    }
}

#[derive(Clone, Debug)]
pub struct ValidationHelp {
    display: bool,
    title: Option<String>,
    text: Option<Box<TextTag>>,
}

impl Default for ValidationHelp {
    fn default() -> Self {
        Self::new()
    }
}

impl ValidationHelp {
    pub fn new() -> Self {
        Self {
            display: true,
            title: None,
            text: None,
        }
    }

    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }

    pub fn display(&self) -> bool {
        self.display
    }

    pub fn set_title(&mut self, title: Option<String>) {
        self.title = title;
    }

    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    pub fn set_text(&mut self, text: Option<TextTag>) {
        if let Some(txt) = text {
            self.text = Some(Box::new(txt));
        } else {
            self.text = None;
        };
    }

    pub fn text(&self) -> Option<&TextTag> {
        self.text.as_deref()
    }
}

#[derive(Copy, Clone, Debug)]
pub enum MessageType {
    Error,
    Warning,
    Info,
}

impl Display for MessageType {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            MessageType::Error => write!(f, "stop"),
            MessageType::Warning => write!(f, "warning"),
            MessageType::Info => write!(f, "information"),
        }
    }
}

#[derive(Clone, Debug)]
pub struct ValidationError {
    display: bool,
    msg_type: MessageType,
    title: Option<String>,
    text: Option<Box<TextTag>>,
}

impl Default for ValidationError {
    fn default() -> Self {
        Self::new()
    }
}

impl ValidationError {
    pub fn new() -> Self {
        Self {
            display: true,
            msg_type: MessageType::Error,
            title: None,
            text: None,
        }
    }

    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }

    pub fn display(&self) -> bool {
        self.display
    }

    pub fn set_msg_type(&mut self, msg_type: MessageType) {
        self.msg_type = msg_type;
    }

    pub fn msg_type(&self) -> &MessageType {
        &self.msg_type
    }

    pub fn set_title(&mut self, title: Option<String>) {
        self.title = title;
    }

    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    pub fn set_text(&mut self, text: Option<TextTag>) {
        if let Some(txt) = text {
            self.text = Some(Box::new(txt));
        } else {
            self.text = None;
        };
    }

    pub fn text(&self) -> Option<&TextTag> {
        self.text.as_deref()
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
    err: Option<ValidationError>,
    help: Option<ValidationHelp>,
}

impl Validation {
    pub fn new() -> Self {
        Self {
            name: "".to_string(),
            condition: "".to_string(),
            base_cell: CellRef::new(),
            allow_empty: true,
            display_list: Default::default(),
            err: Some(ValidationError {
                display: true,
                msg_type: MessageType::Error,
                title: None,
                text: None,
            }),
            help: None,
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
        self.base_cell = base;
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

    /// Error message.
    pub fn set_err(&mut self, err: Option<ValidationError>) {
        self.err = err;
    }

    /// Error message.
    pub fn err(&self) -> Option<&ValidationError> {
        self.err.as_ref()
    }

    /// Help message.
    pub fn set_help(&mut self, help: Option<ValidationHelp>) {
        self.help = help;
    }

    /// Help message.
    pub fn help(&self) -> Option<&ValidationHelp> {
        self.help.as_ref()
    }
}
