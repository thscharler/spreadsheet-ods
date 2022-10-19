//!
//! All kinds of units for use in style attributes.
//!

use crate::OdsError;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

/// Value type for angles.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Angle {
    /// Degrees
    Deg(f64),
    /// Grad degrees.
    Grad(f64),
    /// Radiant.
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

/// Value type for lengths.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Length {
    /// Unspecified length, the actual value is some default or whatever.
    Default,
    /// cm
    Cm(f64),
    /// mm
    Mm(f64),
    /// inch
    In(f64),
    /// typographic points
    Pt(f64),
    /// pica
    Pc(f64),
    /// em
    Em(f64),
}

impl Default for Length {
    fn default() -> Self {
        Length::Default
    }
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
            Length::Default => write!(f, ""),
        }
    }
}

impl FromStr for Length {
    type Err = OdsError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        if s.ends_with("cm") {
            Ok(Length::Cm(s.split_at(s.len() - 2).0.parse()?))
        } else if s.ends_with("mm") {
            Ok(Length::Mm(s.split_at(s.len() - 2).0.parse()?))
        } else if s.ends_with("in") {
            Ok(Length::In(s.split_at(s.len() - 2).0.parse()?))
        } else if s.ends_with("pt") {
            Ok(Length::Pt(s.split_at(s.len() - 2).0.parse()?))
        } else if s.ends_with("pc") {
            Ok(Length::Pc(s.split_at(s.len() - 2).0.parse()?))
        } else if s.ends_with("em") {
            Ok(Length::Em(s.split_at(s.len() - 2).0.parse()?))
        } else {
            Err(OdsError::Parse(s.to_string()))
        }
    }
}

/// Value type that combines lengths and percentages.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Percent {
    /// Percentage
    Percent(f64),
}

impl Display for Percent {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            Percent::Percent(v) => write!(f, "{}%", v),
        }
    }
}

/// Font pitch.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum FontPitch {
    /// Variable font with
    Variable,
    /// Fixed font width
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

/// Various border styles.
#[allow(missing_docs)]
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

/// Page breaks.
#[allow(missing_docs)]
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

/// Line breaks.
#[allow(missing_docs)]
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub enum LineBreak {
    Normal,
    Strict,
}

impl Display for LineBreak {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineBreak::Normal => write!(f, "normal")?,
            LineBreak::Strict => write!(f, "strict")?,
        }
        Ok(())
    }
}

/// Text keep together.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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

/// Writing modes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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

/// Writing modes.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum WritingDirection {
    Ltr,
    Ttb,
}

impl Display for WritingDirection {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            WritingDirection::Ltr => write!(f, "ltr"),
            WritingDirection::Ttb => write!(f, "ttb"),
        }
    }
}

/// Text wrapping.
#[allow(missing_docs)]
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
#[allow(missing_docs)]
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

/// Rotation.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum RotationScale {
    Fixed,
    LineHeight,
}

impl Display for RotationScale {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            RotationScale::Fixed => write!(f, "fixed"),
            RotationScale::LineHeight => write!(f, "line-height"),
        }
    }
}

/// Vertical alignment.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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

/// Fix uses the text-align attribute, value-type bases alignment on content.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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

/// See §7.15.9 of [XSL].
/// If there are no values specified for the fo:text-align and style:justify-single-word
/// 20.301 attributes within the same formatting properties element, the values of those attributes is
/// set to start and false respectively.
///
/// In the OpenDocument XSL-compatible namespace, the fo:text-align attribute does not
/// support the inherit, inside, outside, or string values.
///
/// The fo:text-align attribute is usable with the following elements:
/// <style:list-levelproperties> 17.19 and
/// <style:paragraph-properties> 17.6.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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

/// See §7.15.10 of [XSL].
/// This attribute is ignored if it not accompanied by an fo:text-align 20.223 attribute.
///
/// If no value is specified for this attribute, the value is set to start.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum TextAlignLast {
    Start,
    Center,
    Justify,
}

impl Display for TextAlignLast {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextAlignLast::Start => write!(f, "start"),
            TextAlignLast::Center => write!(f, "center"),
            TextAlignLast::Justify => write!(f, "justify"),
        }
    }
}

/// Vertical alignment for a paragraph.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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

/// Text style values.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FontStyle {
    Normal,
    Italic,
    Oblique,
}

impl Display for FontStyle {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FontStyle::Normal => write!(f, "normal"),
            FontStyle::Italic => write!(f, "italic"),
            FontStyle::Oblique => write!(f, "oblique"),
        }
    }
}

/// Text weight values.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FontWeight {
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

impl Display for FontWeight {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FontWeight::Normal => write!(f, "normal"),
            FontWeight::Bold => write!(f, "bold"),
            FontWeight::W100 => write!(f, "100"),
            FontWeight::W200 => write!(f, "200"),
            FontWeight::W300 => write!(f, "300"),
            FontWeight::W400 => write!(f, "400"),
            FontWeight::W500 => write!(f, "500"),
            FontWeight::W600 => write!(f, "600"),
            FontWeight::W700 => write!(f, "700"),
            FontWeight::W800 => write!(f, "800"),
            FontWeight::W900 => write!(f, "900"),
        }
    }
}

/// Font variant values.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum FontVariant {
    Normal,
    SmallCaps,
}

impl Display for FontVariant {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            FontVariant::Normal => write!(f, "normal"),
            FontVariant::SmallCaps => write!(f, "small-caps"),
        }
    }
}

/// Text case transformations.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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
#[allow(missing_docs)]
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
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum TextPosition {
    Sub,
    Super,
    Percent(Percent),
}

impl Display for TextPosition {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextPosition::Sub => write!(f, "sub"),
            TextPosition::Super => write!(f, "super"),
            TextPosition::Percent(v) => write!(f, "{}", v),
        }
    }
}

/// Text combine.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum TextCombine {
    None,
    Letters,
    Lines,
}

impl Display for TextCombine {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextCombine::None => write!(f, "none"),
            TextCombine::Letters => write!(f, "letters"),
            TextCombine::Lines => write!(f, "lines"),
        }
    }
}

/// Text combine.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum TextEmphasize {
    None,
    Accent,
    Circle,
    Disc,
    Dot,
}

impl Display for TextEmphasize {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            TextEmphasize::None => write!(f, "none"),
            TextEmphasize::Accent => write!(f, "accent"),
            TextEmphasize::Circle => write!(f, "circle"),
            TextEmphasize::Disc => write!(f, "disc"),
            TextEmphasize::Dot => write!(f, "dot"),
        }
    }
}

/// Line style for underline, overline, line-through.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
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
#[allow(missing_docs)]
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
#[allow(missing_docs)]
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
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum LineWidth {
    Auto,
    Bold,
    Percent(Percent),
    Int(u32),
    Length(Length),
    Normal,
    Dash,
    Thin,
    Medium,
    Thick,
}

impl Display for LineWidth {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            LineWidth::Auto => write!(f, "auto"),
            LineWidth::Bold => write!(f, "bold"),
            LineWidth::Percent(v) => write!(f, "{}", v),
            LineWidth::Int(v) => write!(f, "{}", v),
            LineWidth::Length(v) => write!(f, "{}", v),
            LineWidth::Normal => write!(f, "normal"),
            LineWidth::Dash => write!(f, "dash"),
            LineWidth::Thin => write!(f, "thin"),
            LineWidth::Medium => write!(f, "medium"),
            LineWidth::Thick => write!(f, "thick"),
        }
    }
}

/// Page orientation
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum PrintOrientation {
    Landscape,
    Portrait,
}

impl Display for PrintOrientation {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            PrintOrientation::Landscape => write!(f, "landscape"),
            PrintOrientation::Portrait => write!(f, "portrait"),
        }
    }
}

/// The style:cell-protect attribute specifies how a cell is protected.
/// This attribute is only evaluated if the current table is protected.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
#[allow(missing_docs)]
pub enum CellProtect {
    /// If cell content is a formula, it is not displayed. It can be replaced by
    /// changing the cell content.
    /// Note: Replacement of cell content includes replacement with another formula or
    /// other cell content.
    FormulaHidden,
    /// cell content is not displayed and cannot be edited. If content is a
    /// formula, the formula result is not displayed.
    HiddenAndProtected,
    /// Formula responsible for cell content is neither hidden nor protected.
    None,
    /// Cell content cannot be edited.
    Protected,
    /// cell content cannot be edited. If content is a formula, it is not
    /// displayed.
    ProtectedFormulaHidden,
}

impl Display for CellProtect {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            CellProtect::FormulaHidden => write!(f, "formula-hidden"),
            CellProtect::HiddenAndProtected => write!(f, "hidden-and-protected"),
            CellProtect::None => write!(f, "none"),
            CellProtect::Protected => write!(f, "protected"),
            CellProtect::ProtectedFormulaHidden => write!(f, "protected formula-hidden"),
        }
    }
}

/// The style:glyph-orientation-vertical attribute specifies a vertical glyph orientation.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum GlyphOrientation {
    Auto,
    Zero,
    Angle(Angle),
}

impl Display for GlyphOrientation {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            GlyphOrientation::Auto => write!(f, "auto"),
            GlyphOrientation::Zero => write!(f, "0"),
            GlyphOrientation::Angle(a) => a.fmt(f),
        }
    }
}

/// Hyphenation keep. See §7.15.1 of [XSL]
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum Hyphenation {
    Auto,
    Page,
}

impl Display for Hyphenation {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Hyphenation::Auto => write!(f, "auto"),
            Hyphenation::Page => write!(f, "page"),
        }
    }
}

/// Hyphenation ladder count. See §7.15.2 of [XSL].
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum HyphenationLadderCount {
    NoLimit,
    Count(u32),
}

impl Display for HyphenationLadderCount {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            HyphenationLadderCount::NoLimit => write!(f, "no_limit"),
            HyphenationLadderCount::Count(c) => c.fmt(f),
        }
    }
}

/// See §7.15.4 of [XSL](http://www.w3.org/TR/2001/REC-xsl-20011015/).
///
/// The value normal activates the default line height calculation. The value of this attribute can be a
/// length, a percentage, normal.
/// In the OpenDocument XSL-compatible namespace, the fo:line-height attribute does not
/// support the inherit, number, and space values.
///
/// The defined values for the fo:line-height attribute are:
/// * a value of type nonNegativeLength 18.3.20
/// * normal: disables the effects of style:line-height-at-least 20.317 and
/// style:line-spacing 20.318.
/// * a value of type percent 18.3.23
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum LineHeight {
    Normal,
    Length(Length),
    Percent(Percent),
}

impl Display for LineHeight {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            LineHeight::Normal => write!(f, "normal"),
            LineHeight::Length(v) => v.fmt(f),
            LineHeight::Percent(v) => v.fmt(f),
        }
    }
}

/// See §7.29.14 of [XSL].
/// In the OpenDocument XSL-compatible namespace, the fo:margin attribute does not support
/// auto and inherit values.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum Margin {
    Length(Length),
    Percent(Percent),
}

impl Display for Margin {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Margin::Length(v) => v.fmt(f),
            Margin::Percent(v) => v.fmt(f),
        }
    }
}

/// The fo:text-indent attribute specifies a positive or negative indent for the first line of a
/// paragraph. See §7.15.11 of [XSL]. The attribute value is a length. If the attribute is contained in a
/// common style, the attribute value may be also a percentage that refers to the corresponding text
/// indent of a parent style.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum Indent {
    Length(Length),
    Percent(Percent),
}

impl Display for Indent {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Indent::Length(v) => v.fmt(f),
            Indent::Percent(v) => v.fmt(f),
        }
    }
}

/// The style:punctuation-wrap attribute specifies whether a punctuation mark, if one is
/// present, can be hanging, that is, whether it can placed in the margin area at the end of a full line of
/// text.
///
/// The defined values for the style:punctuation-wrap attribute are:
/// * hanging: a punctuation mark can be placed in the margin area at the end of a full line of text.
/// * simple: a punctuation mark cannot be placed in the margin area at the end of a full line of
/// text.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum PunctuationWrap {
    Hanging,
    Simple,
}

impl Display for PunctuationWrap {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            PunctuationWrap::Hanging => write!(f, "hanging"),
            PunctuationWrap::Simple => write!(f, "simple"),
        }
    }
}

/// The style:text-autospace attribute specifies whether to add space between portions of
/// Asian, Western, and complex texts.
///
/// The defined values for the style:text-autospace attribute are:
/// * ideograph-alpha: space should be added between portions of Asian, Western and
/// complex texts.
/// * none: space should not be added between portions of Asian, Western and complex texts.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum TextAutoSpace {
    IdeographAlpha,
    None,
}

impl Display for TextAutoSpace {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            TextAutoSpace::IdeographAlpha => write!(f, "ideograph-alpha"),
            TextAutoSpace::None => write!(f, "none"),
        }
    }
}

/// The style:page-number attribute specifies the page number that should be used for a new
/// page when either a paragraph or table style specifies a master page that should be applied
/// beginning from the start of a paragraph or table.
///
/// The defined values for the style:page-number attribute are:
/// * auto: a page has the page number of the previous page, incremented by one.
/// * a value of type nonNegativeInteger 18.2: specifies a page number.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum PageNumber {
    Auto,
    Number(u32),
}

impl Display for PageNumber {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            PageNumber::Auto => write!(f, "auto"),
            PageNumber::Number(v) => v.fmt(f),
        }
    }
}

/// The style:type attribute specifies the type of a tab stop within paragraph formatting properties.
/// The defined values for the style:type attribute are:
/// * center: text is centered on a tab stop.
/// * char: character appears at a tab stop position.
/// * left: text is left aligned with a tab stop.
/// * right: text is right aligned with a tab stop.
///
/// For a <style:tab-stop> 17.8 element the default value for this attribute is left.
#[derive(Clone, Copy, Debug)]
#[allow(missing_docs)]
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

/// The text:condition attribute specifies the display of text.
/// The defined value of the text:condition attribute is none, which means text is hidden.
/// Works in conjunction with TextDisplay.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum TextCondition {
    None,
}

impl Display for TextCondition {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            TextCondition::None => write!(f, "none"),
        }
    }
}

/// The text:condition attribute specifies the display of text.
/// The defined value of the text:condition attribute is none, which means text is hidden.
/// Works in conjunction with TextDisplay.
#[derive(Debug, Clone, Copy, PartialEq)]
#[allow(missing_docs)]
pub enum TextDisplay {
    None,
    Condition,
    True,
}

impl Display for TextDisplay {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            TextDisplay::None => write!(f, "none"),
            TextDisplay::Condition => write!(f, "condition"),
            TextDisplay::True => write!(f, "true"),
        }
    }
}
