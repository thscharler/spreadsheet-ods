use color::Rgb;
use crate::Style;
use std::fmt::{Display, Formatter};

pub fn font_style(style: &mut Style, ptsize: f32, bold: bool, italic: bool) {
    font_size(style, ptsize);
    font_bold(style, bold);
    font_italic(style, italic);
}

pub fn font_italic(style: &mut Style, italic: bool) {
    if italic {
        style.set_text_prp("fo:font-style", "italic".to_string());
        style.set_text_prp("fo:font-style-asian", "italic".to_string());
        style.set_text_prp("fo:font-style-complex", "italic".to_string());
    } else {
        style.clear_text_prp("fo:font-style");
        style.clear_text_prp("fo:font-style-asian");
        style.clear_text_prp("fo:font-style-complex");
    }
}

pub fn font_bold(style: &mut Style, bold: bool) {
    if bold {
        style.set_text_prp("fo:font-weight", "bold".to_string());
        style.set_text_prp("fo:font-weight-asian", "bold".to_string());
        style.set_text_prp("fo:font-weight-complex", "bold".to_string());
    } else {
        style.clear_text_prp("fo:font-weight");
        style.clear_text_prp("fo:font-weight-asian");
        style.clear_text_prp("fo:font-weight-complex");
    }
}

pub fn font_size(style: &mut Style, ptsize: f32) {
    style.set_text_prp("fo:font-size", format!("{}pt", ptsize));
    style.set_text_prp("fo:font-size-asian", format!("{}pt", ptsize));
    style.set_text_prp("fo:font-size-complex", format!("{}pt", ptsize));
}

pub enum Border {
    None,
    Hidden,
    Dotted,
    Dashed,
    Solid,
    Double,
    Groove,
    Ridge,
    Inset,
    Outset,
}

impl Display for Border {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Border::None => write!(f, "none"),
            Border::Hidden => write!(f, "hidden"),
            Border::Dotted => write!(f, "dotted"),
            Border::Dashed => write!(f, "dashed"),
            Border::Solid => write!(f, "solid"),
            Border::Double => write!(f, "double"),
            Border::Groove => write!(f, "groove"),
            Border::Ridge => write!(f, "ridge"),
            Border::Inset => write!(f, "inset"),
            Border::Outset => write!(f, "outset"),
        }
    }
}

fn border_string(width: f32, border: Border, color: Rgb<u8>) -> String {
    format!("{}pt {} #{:02x}{:02x}{:02x}", width, border, color.r, color.g, color.b)
}

pub fn border_bottom(style: &mut Style, width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-bottom", border_string(width, border, color));
}

pub fn border_top(style: &mut Style, width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-top", border_string(width, border, color));
}

pub fn border_left(style: &mut Style, width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-left", border_string(width, border, color));
}

pub fn border_right(style: &mut Style, width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-right", border_string(width, border, color));
}

pub enum Align {
    Start,
    Center,
    End,
    Justify,
    Inside,
    Outside,
    Left,
    Right,
}

impl Display for Align {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Align::Start => write!(f, "start"),
            Align::Center => write!(f, "center"),
            Align::End => write!(f, "end"),
            Align::Justify => write!(f, "justify"),
            Align::Inside => write!(f, "inside"),
            Align::Outside => write!(f, "outside"),
            Align::Left => write!(f, "left"),
            Align::Right => write!(f, "right"),
        }
    }
}

pub fn text_align(style: &mut Style, align: Align) {
    style.set_paragraph_prp("fo:text-align", align.to_string());
}