use std::collections::HashMap;
use std::fmt::{Display, Formatter};

use color::Rgb;
use string_cache::DefaultAtom;

use crate::{Sheet, ucell, WorkBook, XMLOrigin};
use crate::util::{clear_prp, get_prp, set_prp};

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontDecl {
    pub(crate) name: String,
    /// From where did we get this style.
    pub(crate) origin: XMLOrigin,
    /// All other attributes.
    pub(crate) prp: Option<HashMap<DefaultAtom, String>>,
}

impl FontDecl {
    /// New, empty.
    pub fn new() -> Self {
        FontDecl::new_origin(XMLOrigin::Content)
    }

    /// New, with origination.
    pub fn new_origin(origin: XMLOrigin) -> Self {
        Self {
            name: "".to_string(),
            origin,
            prp: None,
        }
    }

    /// New, with a name.
    pub fn with_name<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            origin: XMLOrigin::Content,
            prp: None,
        }
    }

    ///
    pub fn set_origin(&mut self, origin: XMLOrigin) {
        self.origin = origin;
    }

    pub fn origin(&self) -> XMLOrigin {
        self.origin
    }

    /// Set the name.
    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets a property of the font.
    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    /// Returns a property of the font.
    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }
}

/// Style data fashioned after the ODS spec.
#[derive(Debug, Clone, Default)]
pub struct Style {
    pub(crate) name: String,
    /// Nice String.
    pub(crate) display_name: Option<String>,
    /// From where did we get this style.
    pub(crate) origin: XMLOrigin,
    /// Applicability of this style.
    pub(crate) family: StyleFor,
    /// Styles can cascade.
    pub(crate) parent: Option<String>,
    /// References the actual formatting instructions in the value-styles.
    pub(crate) value_format: Option<String>,
    /// Table styling
    pub(crate) table_prp: Option<HashMap<DefaultAtom, String>>,
    /// Column styling
    pub(crate) table_col_prp: Option<HashMap<DefaultAtom, String>>,
    /// Row styling
    pub(crate) table_row_prp: Option<HashMap<DefaultAtom, String>>,
    /// Cell styles
    pub(crate) table_cell_prp: Option<HashMap<DefaultAtom, String>>,
    /// Cell paragraph styles
    pub(crate) paragraph_prp: Option<HashMap<DefaultAtom, String>>,
    /// Cell text styles
    pub(crate) text_prp: Option<HashMap<DefaultAtom, String>>,
}

impl Style {
    /// New, empty.
    pub fn new() -> Self {
        Style::new_origin(XMLOrigin::Content)
    }

    /// New, with origination.
    pub fn new_origin(origin: XMLOrigin) -> Self {
        Style {
            name: String::from(""),
            display_name: None,
            origin,
            family: StyleFor::None,
            parent: None,
            value_format: None,
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    pub fn cell_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableCell, name, value_style)
    }

    pub fn col_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableColumn, name, value_style)
    }

    pub fn row_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::TableRow, name, value_style)
    }

    pub fn table_style<S: Into<String>, T: Into<String>>(name: S, value_style: T) -> Self {
        Style::with_name(StyleFor::Table, name, value_style)
    }

    /// New, with name.
    pub fn with_name<S: Into<String>, T: Into<String>>(family: StyleFor, name: S, value_style: T) -> Self {
        Style {
            name: name.into(),
            display_name: None,
            origin: XMLOrigin::Content,
            family,
            parent: Some(String::from("Default")),
            value_format: Some(value_style.into()),
            table_prp: None,
            table_col_prp: None,
            table_row_prp: None,
            table_cell_prp: None,
            paragraph_prp: None,
            text_prp: None,
        }
    }

    /// Sets the name.
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Sets the display name.
    pub fn set_display_name<S: Into<String>>(&mut self, name: S) {
        self.display_name = Some(name.into());
    }

    /// Returns the display name.
    pub fn display_name(&self) -> Option<&String> {
        self.display_name.as_ref()
    }

    /// Sets the origin.
    pub fn set_origin(&mut self, origin: XMLOrigin) {
        self.origin = origin;
    }

    /// Returns the origin.
    pub fn origin(&self) -> &XMLOrigin {
        &self.origin
    }

    /// Sets the style-family.
    pub fn set_family(&mut self, family: StyleFor) {
        self.family = family;
    }

    /// Returns the style-family.
    pub fn family(&self) -> &StyleFor {
        &self.family
    }

    /// Sets the parent style.
    pub fn set_parent<S: Into<String>>(&mut self, parent: S) {
        self.parent = Some(parent.into());
    }

    /// Returns the parent style.
    pub fn parent(&self) -> Option<&String> {
        self.parent.as_ref()
    }

    /// Sets the value format.
    pub fn set_value_format<S: Into<String>>(&mut self, value_format: S) {
        self.value_format = Some(value_format.into());
    }

    /// Returns the value format.
    pub fn value_format(&self) -> Option<&String> {
        self.value_format.as_ref()
    }

    /// Sets a property for a table style.
    pub fn set_table_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_prp, name, value);
    }

    /// Returns a property for a table style.
    pub fn table_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_prp, name)
    }

    /// Sets a property for a table column.
    pub fn set_table_col_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_col_prp, name, value);
    }

    /// Returns a property for a table column.
    pub fn table_col_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_col_prp, name)
    }

    /// Set a table row property.
    pub fn set_table_row_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_row_prp, name, value);
    }

    /// Returns a table row property.
    pub fn table_row_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_row_prp, name)
    }

    /// Sets a table cell property.
    pub fn set_table_cell_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.table_cell_prp, name, value);
    }

    /// Returns a table cell property.
    pub fn table_cell_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.table_cell_prp, name)
    }

    /// Sets a text property.
    pub fn set_text_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.text_prp, name, value);
    }

    /// Removes a text property.
    pub fn clear_text_prp(&mut self, name: &str) -> Option<String> {
        clear_prp(&mut self.text_prp, name)
    }

    /// Returns a text property.
    pub fn text_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.text_prp, name)
    }

    /// Sets a paragraph property.
    pub fn set_paragraph_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.paragraph_prp, name, value);
    }

    /// Returns a paragraph property.
    pub fn paragraph_prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.paragraph_prp, name)
    }
}

/// Applicability of this style.
#[derive(Debug, Clone, PartialEq)]
pub enum StyleFor {
    Table,
    TableRow,
    TableColumn,
    TableCell,
    None,
}

impl Default for StyleFor {
    fn default() -> Self {
        StyleFor::None
    }
}


pub fn font_decl<S: Into<String>>(fontdecl: &mut FontDecl, family: S) {
    fontdecl.set_prp("svg:font-family", family.into());
}

pub fn font_name<S: Into<String>>(style: &mut Style, font: S) {
    style.set_text_prp("style:font-name", font.into());
}

pub fn font_style(style: &mut Style, pt_size: f32, bold: bool, italic: bool) {
    font_size(style, pt_size);
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

pub fn font_size(style: &mut Style, pt_size: f32) {
    style.set_text_prp("fo:font-size", format!("{}pt", pt_size));
    style.set_text_prp("fo:font-size-asian", format!("{}pt", pt_size));
    style.set_text_prp("fo:font-size-complex", format!("{}pt", pt_size));
}

pub fn font_color(style: &mut Style, color: Rgb<u8>) {
    style.set_text_prp("fo:color", color_string(color));
}

#[derive(Debug, Clone, Copy)]
pub enum Underline {
    Solid,
    Double,
    Dotted,
    Dashed,
    Wavy,
}

impl Display for Underline {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Underline::Solid => write!(f, "solid"),
            Underline::Double => write!(f, "double"),
            Underline::Dotted => write!(f, "dotted"),
            Underline::Dashed => write!(f, "dashed"),
            Underline::Wavy => write!(f, "wavy"),
        }
    }
}

pub fn font_underline(style: &mut Style, ustyle: Underline) {
    style.set_text_prp("style:text-underline-style", ustyle.to_string());
    style.set_text_prp("style:text-underline-width", "auto".to_string());
    style.set_text_prp("style:text-underline-color", "font-color".to_string());
}

#[derive(Debug, Clone, Copy)]
pub enum LineThroughStyle {
    Dashed,
    DotDash,
    DotDotDash,
    Dotted,
    LongDash,
    None,
    Solid,
    Wave,
}

impl Display for LineThroughStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineThroughStyle::Dashed => write!(f, "dashed"),
            LineThroughStyle::DotDash => write!(f, "dot-dash"),
            LineThroughStyle::DotDotDash => write!(f, "dot-dot-dash"),
            LineThroughStyle::Dotted => write!(f, "dotted"),
            LineThroughStyle::LongDash => write!(f, "long-dash"),
            LineThroughStyle::None => write!(f, "none"),
            LineThroughStyle::Solid => write!(f, "solid"),
            LineThroughStyle::Wave => write!(f, "wavae"),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub enum LineThroughType {
    None,
    Single,
    Double,
}

impl Display for LineThroughType {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineThroughType::None => write!(f, "none"),
            LineThroughType::Single => write!(f, "single"),
            LineThroughType::Double => write!(f, "double"),
        }
    }
}

pub fn font_line_through(style: &mut Style, ltstyle: LineThroughStyle, lttype: LineThroughType) {
    style.set_text_prp("style:text-line-through-style", ltstyle.to_string());
    style.set_text_prp("style:text-line-through-type", lttype.to_string());
}

pub fn font_outline(style: &mut Style, outline: bool) {
    style.set_text_prp("style:text-outline", outline.to_string());
}

pub fn font_shadow(style: &mut Style, pt_shadow_x: f32, pt_shadow_y: f32) {
    style.set_text_prp("fo:text-shadow", format!("{}pt {}pt", pt_shadow_x, pt_shadow_y));
}

#[derive(Debug, Clone, Copy)]
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

fn color_string(color: Rgb<u8>) -> String {
    format!(" #{:02x}{:02x}{:02x}", color.r, color.g, color.b)
}

pub fn cell_border(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border", border_string(pt_width, border, color));
}

pub fn cell_border_bottom(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-bottom", border_string(pt_width, border, color));
}

pub fn cell_border_top(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-top", border_string(pt_width, border, color));
}

pub fn cell_border_left(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-left", border_string(pt_width, border, color));
}

pub fn cell_border_right(style: &mut Style, pt_width: f32, border: Border, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:border-right", border_string(pt_width, border, color));
}

pub fn cell_background(style: &mut Style, color: Rgb<u8>) {
    style.set_table_cell_prp("fo:background-color", color_string(color));
}

pub fn cell_padding(style: &mut Style, pt_padding: f32) {
    style.set_table_cell_prp("fo:padding", format!("{}pt", pt_padding));
}

pub fn cell_shadow(style: &mut Style, pt_off_x: f32, pt_off_y: f32, color: Rgb<u8>) {
    style.set_table_cell_prp("style:shadow", format!("#{:02x}{:02x}{:02x} {}pt {}pt", color.r, color.g, color.b, pt_off_x, pt_off_y));
}

pub fn cell_shrink_to_fit(style: &mut Style, shrink: bool) {
    style.set_table_cell_prp("style:shrink-to-fit", shrink.to_string());
}

pub fn cell_rotation_angle(style: &mut Style, angle: f32) {
    style.set_table_cell_prp("style:rotation-angle", angle.to_string());
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

pub fn set_col_width(workbook: &mut WorkBook, sheet: &mut Sheet, col: ucell, width: &str) {
    let style_name = format!("co{}", col);

    let mut col_style = Style::with_name(StyleFor::TableColumn, &style_name, "");
    col_style.set_table_col_prp("style:column-width", width.to_string());
    workbook.add_style(col_style);

    sheet.set_column_style(col, &style_name);
}

pub fn set_row_height(workbook: &mut WorkBook, sheet: &mut Sheet, row: ucell, height: &str) {
    let style_name = format!("ro{}", row);

    let mut row_style = Style::row_style(&style_name, "");
    row_style.set_table_row_prp("style:row-height", height.to_string());
    row_style.set_table_row_prp("style:use-optimal-row-height", "false".to_string());
    workbook.add_style(row_style);

    sheet.set_row_style(row, &style_name);
}