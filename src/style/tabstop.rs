use crate::attrmap2::AttrMap2;
use crate::style::{color_string, LineStyle, LineType, LineWidth};
use crate::Length;
use color::Rgb;
use std::fmt::{Display, Formatter};

#[derive(Clone, Copy, Debug)]
pub enum TabStopType {
    Center,
    Left,
    Right,
    Char,
}

impl Display for TabStopType {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TabStopType::Center => write!(f, "center"),
            TabStopType::Left => write!(f, "left"),
            TabStopType::Right => write!(f, "right"),
            TabStopType::Char => write!(f, "char"),
        }
    }
}

impl Default for TabStopType {
    fn default() -> Self {
        Self::Left
    }
}

/// Tabstops are part of a paragraph style.
#[derive(Clone, Debug, Default)]
pub struct TabStop {
    attr: AttrMap2,
}

impl TabStop {
    pub fn new() -> Self {
        Self {
            attr: Default::default(),
        }
    }

    /// Delimiter character for tabs of type Char.
    pub fn set_tabstop_char(&mut self, c: char) {
        self.attr.set_attr("style:char", c.to_string());
    }

    /// Color
    pub fn set_leader_color(&mut self, color: Rgb<u8>) {
        self.attr
            .set_attr("style:leader-color", color_string(color));
    }

    /// Linestyle for the leader line.
    pub fn set_leader_style(&mut self, style: LineStyle) {
        self.attr.set_attr("style:leader-style", style.to_string());
    }

    /// Fill character for the leader line.
    pub fn set_leader_text(&mut self, text: char) {
        self.attr.set_attr("style:leader-text", text.to_string());
    }

    /// Textstyle for the leader line.
    pub fn set_leader_text_style(&mut self, styleref: String) {
        self.attr.set_attr("style:leader-text-style", styleref);
    }

    /// LineType for the leader line.
    pub fn set_leader_type(&mut self, t: LineType) {
        self.attr.set_attr("style:leader-type", t.to_string());
    }

    /// Width of the leader line.
    pub fn set_leader_width(&mut self, w: LineWidth) {
        self.attr.set_attr("style:leader-width", w.to_string());
    }

    /// Position of the tab stop.
    pub fn set_position(&mut self, pos: Length) {
        self.attr.set_attr("style:position", pos.to_string());
    }

    /// Type of the tab stop.
    pub fn set_tabstop_type(&mut self, t: TabStopType) {
        self.attr.set_attr("style:type", t.to_string());
    }

    pub fn attr_map(&self) -> &AttrMap2 {
        &self.attr
    }

    pub fn attr_map_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }
}
