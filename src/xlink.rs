//! Enums for XLink.

use get_size2::GetSize;
use std::fmt::{Display, Formatter};

/// See §5.6.2 of XLink.
#[derive(Debug, Clone, Copy, GetSize)]
pub enum XLinkActuate {
    /// XLink
    OnLoad,
    /// XLink
    OnRequest,
}

impl Display for XLinkActuate {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            XLinkActuate::OnLoad => write!(f, "OnLoad"),
            XLinkActuate::OnRequest => write!(f, "OnRequest"),
        }
    }
}

/// See §5.6.1 of XLink.
#[derive(Debug, Clone, Copy, GetSize)]
pub enum XLinkShow {
    /// XLink
    New,
    /// XLink
    Replace,
}

impl Display for XLinkShow {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            XLinkShow::New => write!(f, "new"),
            XLinkShow::Replace => write!(f, "replace"),
        }
    }
}

/// See §3.2 of XLink.
#[derive(Debug, Clone, Copy, Default, GetSize)]
pub enum XLinkType {
    /// XLink
    #[default]
    Simple,
    /// XLink
    Extended,
    /// XLink
    Locator,
    /// XLink
    Arc,
    /// XLink
    Resource,
    /// XLink
    Title,
    /// XLink
    None,
}

impl Display for XLinkType {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            XLinkType::Simple => write!(f, "simple"),
            XLinkType::Extended => write!(f, "extended"),
            XLinkType::Locator => write!(f, "locator"),
            XLinkType::Arc => write!(f, "arc"),
            XLinkType::Resource => write!(f, "resource"),
            XLinkType::Title => write!(f, "title"),
            XLinkType::None => write!(f, "none"),
        }
    }
}
