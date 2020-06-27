//!
//! Defines the type AttrMap as container for different attribute-sets.
//! And there are a number of traits working with AttrMap to set
//! related families of attributes.
//!

use std::collections::{hash_map, HashMap};
use std::fmt::{Display, Formatter};

use color::Rgb;
use string_cache::DefaultAtom;

/// Container type for attributes.
pub type AttrMapType = Option<Box<HashMap<DefaultAtom, String>>>;

/// Container trait for attributes.
pub trait AttrMap {
    /// Reference to the map of actual attributes.
    fn attr_map(&self) -> &AttrMapType;
    /// Reference to the map of actual attributes.
    fn attr_map_mut(&mut self) -> &mut AttrMapType;

    /// Are there any attributes?
    fn has_attr(&self) -> bool {
        self.attr_map().is_none()
    }

    /// Add from Vec
    fn add_all(&mut self, data: Vec<(&str, String)>) {
        let attr = self.attr_map_mut();

        let attr = attr.get_or_insert_with(|| Box::new(HashMap::new()));
        for (name, val) in data {
            attr.insert(DefaultAtom::from(name), val);
        }
    }

    /// Adds an attribute.
    fn set_attr(&mut self, name: &str, value: String) {
        self.attr_map_mut()
            .get_or_insert_with(|| Box::new(HashMap::new()))
            .insert(DefaultAtom::from(name), value);
    }

    /// Removes an attribute.
    fn clear_attr(&mut self, name: &str) -> Option<String> {
        if let Some(ref mut attr) = self.attr_map_mut() {
            attr.remove(&DefaultAtom::from(name))
        } else {
            None
        }
    }

    /// Returns the attribute.
    fn attr(&self, name: &str) -> Option<&String> {
        if let Some(prp) = self.attr_map() {
            prp.get(&DefaultAtom::from(name))
        } else {
            None
        }
    }
}

/// Iterator for an AttrMap.
pub struct AttrMapIter<'a> {
    it: Option<hash_map::Iter<'a, DefaultAtom, String>>,
}

impl<'a> From<&'a AttrMapType> for AttrMapIter<'a> {
    fn from(attrmap: &'a AttrMapType) -> Self {
        if let Some(attrmap) = attrmap {
            Self {
                it: Some(attrmap.iter()),
            }
        } else {
            Self { it: None }
        }
    }
}

impl<'a> Iterator for AttrMapIter<'a> {
    type Item = (&'a DefaultAtom, &'a String);

    fn next(&mut self) -> Option<Self::Item> {
        if let Some(it) = &mut self.it {
            it.next()
        } else {
            None
        }
    }
}

/// Value type for angles.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Angle {
    Deg(f64),
    Grad(f64),
    Rad(f64),
}

impl Display for Angle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Angle::Deg(v) => write!(f, "{}deg", v),
            Angle::Grad(v) => write!(f, "{}grad", v),
            Angle::Rad(v) => write!(f, "{}rad", v),
        }
    }
}

/// deg angles. 360°
#[macro_export]
macro_rules! deg {
    ($l:expr) => {
        Angle::Deg($l as f64)
    };
}

/// grad angles. 400°
#[macro_export]
macro_rules! grad {
    ($l:expr) => {
        Length::Grad($l as f64)
    };
}

/// radians angle.
#[macro_export]
macro_rules! rad {
    ($l:expr) => {
        Length::Rad($l as f64)
    };
}

/// Value type for lengths.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Length {
    Cm(f64),
    Mm(f64),
    In(f64),
    Pt(f64),
    Pc(f64),
    Em(f64),
}

impl Display for Length {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Length::Cm(v) => write!(f, "{}cm", v),
            Length::Mm(v) => write!(f, "{}mm", v),
            Length::In(v) => write!(f, "{}in", v),
            Length::Pt(v) => write!(f, "{}pt", v),
            Length::Pc(v) => write!(f, "{}pc", v),
            Length::Em(v) => write!(f, "{}em", v),
        }
    }
}

/// Centimeters.
#[macro_export]
macro_rules! cm {
    ($l:expr) => {
        Length::Cm($l as f64)
    };
}

/// Millimeters.
#[macro_export]
macro_rules! mm {
    ($l:expr) => {
        Length::Mm($l as f64)
    };
}

/// Inches.
#[macro_export]
macro_rules! inch {
    ($l:expr) => {
        Length::In($l as f64)
    };
}

/// Point. 1/72"
#[macro_export]
macro_rules! pt {
    ($l:expr) => {
        Length::Pt($l as f64)
    };
}

/// Pica. 12/72"
#[macro_export]
macro_rules! pc {
    ($l:expr) => {
        Length::Pc($l as f64)
    };
}

/// Length depending on font size.
#[macro_export]
macro_rules! em {
    ($l:expr) => {
        Length::Em($l as f64)
    };
}

/// Font pitch.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum FontPitch {
    Variable,
    Fixed,
}

impl Display for FontPitch {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FontPitch::Variable => write!(f, "variable"),
            FontPitch::Fixed => write!(f, "fixed"),
        }
    }
}

/// Attributes for a FontFaceDecl
pub trait AttrFontDecl
where
    Self: AttrMap,
{
    /// Font-name
    fn set_name<S: Into<String>>(&mut self, name: S) {
        self.set_attr("style:name", name.into());
    }

    /// External font family name.
    fn set_font_family<S: Into<String>>(&mut self, name: S) {
        self.set_attr("svg:font-family", name.into());
    }

    /// System generic name.
    fn set_font_family_generic<S: Into<String>>(&mut self, name: S) {
        self.set_attr("style:font-family-generic", name.into());
    }

    /// Font pitch.
    fn set_font_pitch(&mut self, pitch: FontPitch) {
        self.set_attr("style:font-pitch", pitch.to_string());
    }
}

/// Margin attributes.
pub trait AttrFoMargin
where
    Self: AttrMap,
{
    fn set_margin(&mut self, margin: Length) {
        self.set_attr("fo:margin", margin.to_string());
    }

    fn set_margin_bottom(&mut self, margin: Length) {
        self.set_attr("fo:margin-bottom", margin.to_string());
    }

    fn set_margin_left(&mut self, margin: Length) {
        self.set_attr("fo:margin-left", margin.to_string());
    }

    fn set_margin_right(&mut self, margin: Length) {
        self.set_attr("fo:margin-right", margin.to_string());
    }

    fn set_margin_top(&mut self, margin: Length) {
        self.set_attr("fo:margin-top", margin.to_string());
    }
}

/// Padding attributes.
pub trait AttrFoPadding
where
    Self: AttrMap,
{
    fn set_padding(&mut self, padding: Length) {
        self.set_attr("fo:padding", padding.to_string());
    }

    fn set_padding_bottom(&mut self, padding: Length) {
        self.set_attr("fo:padding-bottom", padding.to_string());
    }

    fn set_padding_left(&mut self, padding: Length) {
        self.set_attr("fo:padding-left", padding.to_string());
    }

    fn set_padding_right(&mut self, padding: Length) {
        self.set_attr("fo:padding-right", padding.to_string());
    }

    fn set_padding_top(&mut self, padding: Length) {
        self.set_attr("fo:padding-top", padding.to_string());
    }
}

/// Background color.
pub trait AttrFoBackgroundColor
where
    Self: AttrMap,
{
    /// Border style.
    fn set_background_color(&mut self, color: Rgb<u8>) {
        self.set_attr("fo:background-color", color_string(color));
    }
}

/// Minimum height.
pub trait AttrFoMinHeight
where
    Self: AttrMap,
{
    fn set_min_height(&mut self, height: Length) {
        self.set_attr("fo:min-height", height.to_string());
    }

    fn set_min_height_percent(&mut self, height: f64) {
        self.set_attr("fo:min-height", percent_string(height));
    }
}

/// Various border styles.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
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

/// Border attributes.
pub trait AttrFoBorder
where
    Self: AttrMap,
{
    /// Border style all four sides.
    fn set_border(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_bottom(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-bottom", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_top(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-top", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_left(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-left", border_string(width, border, color));
    }

    /// Border style.
    fn set_border_right(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-right", border_string(width, border, color));
    }

    /// Widths for double borders.
    fn set_border_line_width(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_bottom(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-bottom",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_left(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-left",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_right(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-right",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Widths for double borders.
    fn set_border_line_width_top(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:border-line-width-top",
            border_line_width_string(inner, spacing, outer),
        );
    }
}

/// Page breaks.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum PageBreak {
    Auto,
    Column,
    Page,
}

impl Display for PageBreak {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            PageBreak::Auto => write!(f, "auto")?,
            PageBreak::Column => write!(f, "column")?,
            PageBreak::Page => write!(f, "page")?,
        }
        Ok(())
    }
}

/// Page breaks.
pub trait AttrFoBreak
where
    Self: AttrMap,
{
    /// page-break
    fn set_break_before(&mut self, pagebreak: PageBreak) {
        self.set_attr("fo:break-before", pagebreak.to_string());
    }

    // page-break
    fn set_break_after(&mut self, pagebreak: PageBreak) {
        self.set_attr("fo:break-after", pagebreak.to_string());
    }
}

/// Text keep together.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextKeep {
    Auto,
    Always,
}

impl Display for TextKeep {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextKeep::Auto => write!(f, "auto")?,
            TextKeep::Always => write!(f, "always")?,
        }
        Ok(())
    }
}

/// Keep with next.
pub trait AttrFoKeepWithNext
where
    Self: AttrMap,
{
    /// page-break
    fn set_keep_with_next(&mut self, keep_with_next: TextKeep) {
        self.set_attr("fo:keep-with-next", keep_with_next.to_string());
    }
}

/// Keep together.
pub trait AttrFoKeepTogether
where
    Self: AttrMap,
{
    /// page-break
    fn set_keep_together(&mut self, keep_together: TextKeep) {
        self.set_attr("fo:keep-together", keep_together.to_string());
    }
}

/// Height attribute.
pub trait AttrSvgHeight
where
    Self: AttrMap,
{
    fn set_height(&mut self, height: Length) {
        self.set_attr("svg:height", height.to_string());
    }
}

/// Spacing for header/footer.
pub trait AttrStyleDynamicSpacing
where
    Self: AttrMap,
{
    fn set_dynamic_spacing(&mut self, dynamic: bool) {
        self.set_attr("style:dynamic-spacing", dynamic.to_string());
    }
}

/// Shadows. Only a single shadow supported here.
pub trait AttrStyleShadow
where
    Self: AttrMap,
{
    fn set_shadow(
        &mut self,
        x_offset: Length,
        y_offset: Length,
        blur: Option<Length>,
        color: Rgb<u8>,
    ) {
        self.set_attr(
            "style:shadow",
            shadow_string(x_offset, y_offset, blur, color),
        );
    }
}

/// Writing modes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum WritingMode {
    LrTb,
    RlTb,
    TbRl,
    TbLr,
    Lr,
    Rl,
    Tb,
    Page,
}

impl Display for WritingMode {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            WritingMode::LrTb => write!(f, "lr-tb"),
            WritingMode::RlTb => write!(f, "rl-tb"),
            WritingMode::TbRl => write!(f, "tb-rl"),
            WritingMode::TbLr => write!(f, "tb-lr"),
            WritingMode::Lr => write!(f, "lr"),
            WritingMode::Rl => write!(f, "rl"),
            WritingMode::Tb => write!(f, "tb"),
            WritingMode::Page => write!(f, "page"),
        }
    }
}

/// Writing mode.
pub trait AttrStyleWritingMode
where
    Self: AttrMap,
{
    fn set_writing_mode(&mut self, writing_mode: WritingMode) {
        self.set_attr("style:writing-mode", writing_mode.to_string());
    }
}

/// Table row specific attributes.
pub trait AttrTableRow
where
    Self: AttrMap,
{
    fn set_min_row_height(&mut self, min_height: Length) {
        self.set_attr("style:min-row-height", min_height.to_string());
    }

    fn set_row_height(&mut self, height: Length) {
        self.set_attr("style:row-height", height.to_string());
    }

    fn set_use_optimal_row_height(&mut self, opt: bool) {
        self.set_attr("style:use-optimal-row-height", opt.to_string());
    }
}

/// Table columns specific attributes.
pub trait AttrTableCol
where
    Self: AttrMap,
{
    /// Relative weights for the column width
    fn set_rel_col_width(&mut self, rel: f64) {
        self.set_attr("style:rel-column-width", rel_width_string(rel));
    }

    /// Column width
    fn set_col_width(&mut self, width: Length) {
        self.set_attr("style:column-width", width.to_string());
    }

    /// Override switch for the column width.
    fn set_use_optimal_col_width(&mut self, opt: bool) {
        self.set_attr("style:use-optimal-column-width", opt.to_string());
    }
}

/// Text wrapping.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum WrapOption {
    NoWrap,
    Wrap,
}

impl Display for WrapOption {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            WrapOption::NoWrap => write!(f, "no-wrap"),
            WrapOption::Wrap => write!(f, "wrap"),
        }
    }
}

/// Rotation.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum RotationAlign {
    None,
    Bottom,
    Top,
    Center,
}

impl Display for RotationAlign {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            RotationAlign::None => write!(f, "none"),
            RotationAlign::Bottom => write!(f, "bottom"),
            RotationAlign::Top => write!(f, "top"),
            RotationAlign::Center => write!(f, "center"),
        }
    }
}

/// Vertical alignment.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum CellAlignVertical {
    Top,
    Middle,
    Bottom,
    Automatic,
}

impl Display for CellAlignVertical {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            CellAlignVertical::Top => write!(f, "top"),
            CellAlignVertical::Middle => write!(f, "middle"),
            CellAlignVertical::Bottom => write!(f, "bottom"),
            CellAlignVertical::Automatic => write!(f, "automatic"),
        }
    }
}

/// Table cell specific styles.
pub trait AttrTableCell
where
    Self: AttrMap,
{
    fn set_wrap_option(&mut self, wrap: WrapOption) {
        self.set_attr("fo:wrap-option", wrap.to_string());
    }

    fn set_print_content(&mut self, print: bool) {
        self.set_attr("style:print-content", print.to_string());
    }

    fn set_repeat_content(&mut self, print: bool) {
        self.set_attr("style:repeat-content", print.to_string());
    }

    fn set_rotation_align(&mut self, align: RotationAlign) {
        self.set_attr("style:rotation-align", align.to_string());
    }

    fn set_rotation_angle(&mut self, angle: Angle) {
        self.set_attr("style:rotation-angle", angle.to_string());
    }

    fn set_shrink_to_fit(&mut self, shrink: bool) {
        self.set_attr("style:shrink-to-fit", shrink.to_string());
    }

    fn set_vertical_align(&mut self, align: CellAlignVertical) {
        self.set_attr("style:vertical-align", align.to_string());
    }

    /// Diagonal style.
    fn set_diagonal_bl_tr(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("style:diagonal-bl-tr", border_string(width, border, color));
    }

    /// Widths for double borders.
    fn set_diagonal_bl_tr_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:diagonal-bl-tr-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }

    /// Diagonal style.
    fn set_diagonal_tl_br(&mut self, width: Length, border: Border, color: Rgb<u8>) {
        self.set_attr("style:diagonal-tl-br", border_string(width, border, color));
    }

    /// Widths for double borders.
    fn set_diagonal_tl_br_widths(&mut self, inner: Length, spacing: Length, outer: Length) {
        self.set_attr(
            "style:diagonal-tl-br-widths",
            border_line_width_string(inner, spacing, outer),
        );
    }
}

/// Fix uses the text-align attribute, value-type bases alignment on content.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextAlignSource {
    Fix,
    ValueType,
}

impl Display for TextAlignSource {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextAlignSource::Fix => write!(f, "fix"),
            TextAlignSource::ValueType => write!(f, "value-type"),
        }
    }
}

/// Horizontal alignment.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextAlign {
    Start,
    Center,
    End,
    Justify,
    Inside,
    Outside,
    Left,
    Right,
}

impl Display for TextAlign {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextAlign::Start => write!(f, "start"),
            TextAlign::Center => write!(f, "center"),
            TextAlign::End => write!(f, "end"),
            TextAlign::Justify => write!(f, "justify"),
            TextAlign::Inside => write!(f, "inside"),
            TextAlign::Outside => write!(f, "outside"),
            TextAlign::Left => write!(f, "left"),
            TextAlign::Right => write!(f, "right"),
        }
    }
}

/// Vertical alignment.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum ParaAlignVertical {
    Top,
    Middle,
    Bottom,
    Auto,
    Baseline,
}

impl Display for ParaAlignVertical {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            ParaAlignVertical::Top => write!(f, "top"),
            ParaAlignVertical::Middle => write!(f, "middle"),
            ParaAlignVertical::Bottom => write!(f, "bottom"),
            ParaAlignVertical::Auto => write!(f, "auto"),
            ParaAlignVertical::Baseline => write!(f, "baseline"),
        }
    }
}

/// Paragraph specific styles.
pub trait AttrParagraph
where
    Self: AttrMap,
{
    fn set_text_align_source(&mut self, align: TextAlignSource) {
        self.set_attr("style:text-align-source", align.to_string());
    }

    fn set_text_align(&mut self, align: TextAlign) {
        self.set_attr("fo:text-align", align.to_string());
    }

    fn set_text_indent(&mut self, indent: Length) {
        self.set_attr("fo:text-indent", indent.to_string());
    }

    fn set_line_spacing(&mut self, spacing: Length) {
        self.set_attr("style:line-spacing", spacing.to_string());
    }

    fn set_number_lines(&mut self, number: bool) {
        self.set_attr("text:number-lines", number.to_string());
    }

    fn set_vertical_align(&mut self, align: ParaAlignVertical) {
        self.set_attr("style:vertical-align", align.to_string());
    }
}

/// Text style values.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextStyle {
    Normal,
    Italic,
    Oblique,
}

impl Display for TextStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextStyle::Normal => write!(f, "normal"),
            TextStyle::Italic => write!(f, "italic"),
            TextStyle::Oblique => write!(f, "oblique"),
        }
    }
}

/// Text weight values.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextWeight {
    Normal,
    Bold,
    W100,
    W200,
    W300,
    W400,
    W500,
    W600,
    W700,
    W800,
    W900,
}

impl Display for TextWeight {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextWeight::Normal => write!(f, "normal"),
            TextWeight::Bold => write!(f, "bold"),
            TextWeight::W100 => write!(f, "100"),
            TextWeight::W200 => write!(f, "200"),
            TextWeight::W300 => write!(f, "300"),
            TextWeight::W400 => write!(f, "400"),
            TextWeight::W500 => write!(f, "500"),
            TextWeight::W600 => write!(f, "600"),
            TextWeight::W700 => write!(f, "700"),
            TextWeight::W800 => write!(f, "800"),
            TextWeight::W900 => write!(f, "900"),
        }
    }
}

/// Text case transformations.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextTransform {
    None,
    Lowercase,
    Uppercase,
    Capitalize,
}

impl Display for TextTransform {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextTransform::None => write!(f, "none"),
            TextTransform::Lowercase => write!(f, "lowercase"),
            TextTransform::Uppercase => write!(f, "uppercase"),
            TextTransform::Capitalize => write!(f, "capitalize"),
        }
    }
}

/// Text style engraved and embossed.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextRelief {
    None,
    Embossed,
    Engraved,
}

impl Display for TextRelief {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextRelief::None => write!(f, "none"),
            TextRelief::Embossed => write!(f, "embossed"),
            TextRelief::Engraved => write!(f, "engraved"),
        }
    }
}

/// Text style subscript or superscript.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum TextPosition {
    Sub,
    Super,
}

impl Display for TextPosition {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextPosition::Sub => write!(f, "sub"),
            TextPosition::Super => write!(f, "super"),
        }
    }
}

/// Line style for underline, overline, line-through.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum LineStyle {
    Dash,
    DotDash,
    DotDotDash,
    Dotted,
    LongDash,
    None,
    Solid,
    Wave,
}

impl Display for LineStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineStyle::Dash => write!(f, "dash"),
            LineStyle::DotDash => write!(f, "dot-dash"),
            LineStyle::DotDotDash => write!(f, "dot-dot-dash"),
            LineStyle::Dotted => write!(f, "dotted"),
            LineStyle::LongDash => write!(f, "long-dash"),
            LineStyle::None => write!(f, "none"),
            LineStyle::Solid => write!(f, "solid"),
            LineStyle::Wave => write!(f, "wave"),
        }
    }
}

/// Line types for underline, overline, line-through.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum LineType {
    None,
    Single,
    Double,
}

impl Display for LineType {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineType::None => write!(f, "none"),
            LineType::Single => write!(f, "single"),
            LineType::Double => write!(f, "double"),
        }
    }
}

/// Line modes for underline, overline, line-through.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum LineMode {
    Continuous,
    SkipWhiteSpace,
}

impl Display for LineMode {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineMode::Continuous => write!(f, "continuous"),
            LineMode::SkipWhiteSpace => write!(f, "skip-white-space"),
        }
    }
}

/// Line width for underline, overline, line-through.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum LineWidth {
    Auto,
    Normal,
    Bold,
    Thin,
    Medium,
    Thick,
}

impl Display for LineWidth {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineWidth::Auto => write!(f, "auto"),
            LineWidth::Normal => write!(f, "normal"),
            LineWidth::Bold => write!(f, "bold"),
            LineWidth::Thin => write!(f, "thin"),
            LineWidth::Medium => write!(f, "medium"),
            LineWidth::Thick => write!(f, "thick"),
        }
    }
}

/// Text style attributes.
pub trait AttrText
where
    Self: AttrMap,
{
    fn set_color(&mut self, color: Rgb<u8>) {
        self.set_attr("fo:color", color_string(color));
    }

    fn set_font_name<S: Into<String>>(&mut self, name: S) {
        self.set_attr("style:font-name", name.into());
    }

    fn set_font_attr(&mut self, size: Length, bold: bool, italic: bool) {
        self.set_font_size(size);
        if bold {
            self.set_font_italic();
        }
        if italic {
            self.set_font_bold();
        }
    }

    fn set_font_size(&mut self, size: Length) {
        self.set_attr("fo:font-size", size.to_string());
    }

    fn set_font_size_percent(&mut self, size: f64) {
        self.set_attr("fo:font-size", percent_string(size));
    }

    fn set_font_italic(&mut self) {
        self.set_attr("fo:font-style", "italic".to_string());
    }

    fn set_font_style(&mut self, style: TextStyle) {
        self.set_attr("fo:font-style", style.to_string());
    }

    fn set_font_bold(&mut self) {
        self.set_attr("fo:font-weight", TextWeight::Bold.to_string());
    }

    fn set_font_weight(&mut self, weight: TextWeight) {
        self.set_attr("fo:font-weight", weight.to_string());
    }

    fn set_letter_spacing(&mut self, spacing: Length) {
        self.set_attr("fo:letter-spacing", spacing.to_string());
    }

    fn set_letter_spacing_normal(&mut self) {
        self.set_attr("fo:letter-spacing", "normal".to_string());
    }

    fn set_text_shadow(
        &mut self,
        x_offset: Length,
        y_offset: Length,
        blur: Option<Length>,
        color: Rgb<u8>,
    ) {
        self.set_attr(
            "fo:text-shadow",
            shadow_string(x_offset, y_offset, blur, color),
        );
    }

    fn set_text_position(&mut self, pos: TextPosition) {
        self.set_attr("style:text-position", pos.to_string());
    }

    fn set_text_transform(&mut self, trans: TextTransform) {
        self.set_attr("fo:text-transform", trans.to_string());
    }

    fn set_font_relief(&mut self, relief: TextRelief) {
        self.set_attr("style:font-relief", relief.to_string());
    }

    fn set_font_line_through_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:text-line-through-color", color_string(color));
    }

    fn set_font_line_through_style(&mut self, lstyle: LineStyle) {
        self.set_attr("style:text-line-through-style", lstyle.to_string());
    }

    fn set_font_line_through_mode(&mut self, lmode: LineMode) {
        self.set_attr("style:text-line-through-mode", lmode.to_string());
    }

    fn set_font_line_through_type(&mut self, ltype: LineType) {
        self.set_attr("style:text-line-through-type", ltype.to_string());
    }

    fn set_font_line_through_text<S: Into<String>>(&mut self, text: S) {
        self.set_attr("style:text-line-through-text", text.into());
    }

    fn set_font_line_through_text_style<S: Into<String>>(&mut self, style_ref: S) {
        self.set_attr("style:text-line-through-text-style", style_ref.into());
    }

    fn set_font_line_through_width(&mut self, lwidth: LineWidth) {
        self.set_attr("style:text-line-through-width", lwidth.to_string());
    }

    fn set_font_text_outline(&mut self, outline: bool) {
        self.set_attr("style:text-outline", outline.to_string());
    }

    fn set_font_underline_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:text-underline-color", color_string(color));
    }

    fn set_font_underline_style(&mut self, lstyle: LineStyle) {
        self.set_attr("style:text-underline-style", lstyle.to_string());
    }

    fn set_font_underline_type(&mut self, ltype: LineType) {
        self.set_attr("style:text-underline-type", ltype.to_string());
    }

    fn set_font_underline_mode(&mut self, lmode: LineMode) {
        self.set_attr("style:text-underline-mode", lmode.to_string());
    }

    fn set_font_underline_width(&mut self, lwidth: LineWidth) {
        self.set_attr("style:text-underline-width", lwidth.to_string());
    }

    fn set_font_overline_color(&mut self, color: Rgb<u8>) {
        self.set_attr("style:text-overline-color", color_string(color));
    }

    fn set_font_overline_style(&mut self, lstyle: LineStyle) {
        self.set_attr("style:text-overline-style", lstyle.to_string());
    }

    fn set_font_overline_type(&mut self, ltype: LineType) {
        self.set_attr("style:text-overline-type", ltype.to_string());
    }

    fn set_font_overline_mode(&mut self, lmode: LineMode) {
        self.set_attr("style:text-overline-mode", lmode.to_string());
    }

    fn set_font_overline_width(&mut self, lwidth: LineWidth) {
        self.set_attr("style:text-overline-width", lwidth.to_string());
    }
}

pub(crate) fn color_string(color: Rgb<u8>) -> String {
    format!("#{:02x}{:02x}{:02x}", color.r, color.g, color.b)
}

fn border_string(width: Length, border: Border, color: Rgb<u8>) -> String {
    format!(
        "{} {} #{:02x}{:02x}{:02x}",
        width, border, color.r, color.g, color.b
    )
}

fn percent_string(value: f64) -> String {
    format!("{}%", value)
}

fn rel_width_string(value: f64) -> String {
    format!("{}*", value)
}

fn border_line_width_string(inner: Length, space: Length, outer: Length) -> String {
    format!("{} {} {}", inner, space, outer)
}

fn shadow_string(
    x_offset: Length,
    y_offset: Length,
    blur: Option<Length>,
    color: Rgb<u8>,
) -> String {
    if let Some(blur) = blur {
        format!("{} {} {} {}", color_string(color), x_offset, y_offset, blur)
    } else {
        format!("{} {} {}", color_string(color), x_offset, y_offset)
    }
}
