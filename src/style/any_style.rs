use crate::style::{StyleOrigin, StyleUse, TableStyle};

#[derive(Debug, Clone)]
pub enum AnyStyle {
    TableStyle(TableStyle),
}

impl AnyStyle {
    pub fn origin(&self) -> StyleOrigin {
        match self {
            AnyStyle::TableStyle(s) => s.origin(),
        }
    }

    pub fn styleuse(&self) -> StyleUse {
        match self {
            AnyStyle::TableStyle(s) => s.styleuse(),
        }
    }

    pub fn name(&self) -> &str {
        let name = match self {
            AnyStyle::TableStyle(s) => s.name(),
        };

        if let Some(name) = name {
            name
        } else {
            ""
        }
    }
}
