use std::collections::HashMap;
use std::fmt::{Display, Formatter};

use chrono::NaiveDateTime;
use string_cache::DefaultAtom;
use time::Duration;

use crate::{StyleOrigin, StyleUse, ValueType};
use crate::util::{get_prp, get_prp_def, set_prp, set_prp_vec};

#[derive(Debug)]
pub enum ValueFormatError {
    Format(String),
    NaN,
}

impl Display for ValueFormatError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
        match self {
            ValueFormatError::Format(s) => write!(f, "{}", s)?,
            ValueFormatError::NaN => write!(f, "Digit expected")?,
        }
        Ok(())
    }
}

impl std::error::Error for ValueFormatError {}

/// Actual textual formatting of values.
#[derive(Debug, Clone, Default)]
pub struct ValueFormat {
    // Name
    pub(crate) name: String,
    // Value type
    pub(crate) v_type: ValueType,
    // Origin information.
    pub(crate) origin: StyleOrigin,
    // Usage of this style.
    pub(crate) styleuse: StyleUse,
    // Properties of the format.
    pub(crate) prp: Option<HashMap<DefaultAtom, String>>,
    // Parts of the format.
    pub(crate) parts: Option<Vec<FormatPart>>,
}

impl ValueFormat {
    /// New, empty.
    pub fn new() -> Self {
        ValueFormat::new_origin(Default::default(), Default::default())
    }

    /// New, with origin.
    pub fn new_origin(origin: StyleOrigin, styleuse: StyleUse) -> Self {
        ValueFormat {
            name: String::from(""),
            v_type: ValueType::Text,
            origin,
            styleuse,
            prp: None,
            parts: None,
        }
    }

    /// New, with name.
    pub fn with_name<S: Into<String>>(name: S, value_type: ValueType) -> Self {
        ValueFormat {
            name: name.into(),
            v_type: value_type,
            origin: Default::default(),
            styleuse: Default::default(),
            prp: None,
            parts: None,
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

    /// Sets the value type.
    pub fn set_value_type(&mut self, value_type: ValueType) {
        self.v_type = value_type;
    }

    /// Returns the value type.
    pub fn value_type(&self) -> ValueType {
        self.v_type
    }

    /// Sets the origin.
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Returns the origin.
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// Style usage.
    pub fn set_styleuse(&mut self, styleuse: StyleUse) {
        self.styleuse = styleuse;
    }

    /// Returns the usage.
    pub fn styleuse(&self) -> StyleUse {
        self.styleuse
    }

    /// Sets a property of the format.
    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    /// Returns a property of the format.
    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    /// Adds a format part.
    pub fn push_part(&mut self, part: FormatPart) {
        if let Some(parts) = &mut self.parts {
            parts.push(part);
        } else {
            self.parts = Some(vec![part]);
        }
    }

    /// Adds all format parts.
    pub fn push_parts(&mut self, parts: Vec<FormatPart>) {
        for p in parts.into_iter() {
            self.push_part(p);
        }
    }

    /// Returns the parts.
    pub fn parts(&self) -> Option<&Vec<FormatPart>> {
        self.parts.as_ref()
    }

    /// Returns the mutable parts.
    pub fn parts_mut(&mut self) -> &mut Vec<FormatPart> {
        self.parts.get_or_insert(Vec::new())
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    pub fn format_boolean(&self, b: bool) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_boolean(&mut buf, b);
            }
        }
        buf
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    pub fn format_float(&self, f: f64) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_float(&mut buf, f);
            }
        }
        buf
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    pub fn format_str(&self, s: &str) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_str(&mut buf, s);
            }
        }
        buf
    }

    // Tries to format.
    // If there are no matching parts, does nothing.
    // Should work reasonably. Don't ask me about other calenders.
    pub fn format_datetime(&self, d: &NaiveDateTime) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            let h12 = parts.iter().any(|v| v.part_type == FormatPartType::AmPm);

            for p in parts {
                p.format_datetime(&mut buf, d, h12);
            }
        }
        buf
    }

    // Tries to format. Should work reasonably.
    // If there are no matching parts, does nothing.
    pub fn format_time_duration(&self, d: &Duration) -> String {
        let mut buf = String::new();
        if let Some(parts) = &self.parts {
            for p in parts {
                p.format_time_duration(&mut buf, d);
            }
        }
        buf
    }
}

/// Identifies the structural parts of a value format.
#[derive(Debug, Clone, PartialEq)]
pub enum FormatPartType {
    Boolean,
    Number,
    Fraction,
    Scientific,
    CurrencySymbol,
    Day,
    Month,
    Year,
    Era,
    DayOfWeek,
    WeekOfYear,
    Quarter,
    Hours,
    Minutes,
    Seconds,
    AmPm,
    EmbeddedText,
    Text,
    TextContent,
    StyleText,
    StyleMap,
}

/// One structural part of a value format.
#[derive(Debug, Clone)]
pub struct FormatPart {
    // What kind of format part is this?
    pub(crate) part_type: FormatPartType,
    // Properties of this part.
    pub(crate) prp: Option<HashMap<DefaultAtom, String>>,
    // Some content.
    pub(crate) content: Option<String>,
}

impl FormatPart {
    /// New, empty
    pub fn new(ftype: FormatPartType) -> Self {
        FormatPart {
            part_type: ftype,
            prp: None,
            content: None,
        }
    }

    /// New, with string content.
    pub fn new_content(ftype: FormatPartType, content: &str) -> Self {
        FormatPart {
            part_type: ftype,
            prp: None,
            content: Some(content.to_string()),
        }
    }

    /// New with properties.
    pub fn new_vec(ftype: FormatPartType, prp_vec: Vec<(&str, String)>) -> Self {
        let mut part = FormatPart {
            part_type: ftype,
            prp: None,
            content: None,
        };
        part.set_prp_vec(prp_vec);
        part
    }

    /// Sets the kind of the part.
    pub fn set_part_type(&mut self, p_type: FormatPartType) {
        self.part_type = p_type;
    }

    /// What kind of part?
    pub fn part_type(&self) -> &FormatPartType {
        &self.part_type
    }

    /// Sets a vec of properties.
    pub fn set_prp_vec(&mut self, vec: Vec<(&str, String)>) {
        set_prp_vec(&mut self.prp, vec);
    }

    /// Sets a property.
    pub fn set_prp(&mut self, name: &str, value: String) {
        set_prp(&mut self.prp, name, value);
    }

    /// Returns a property.
    pub fn prp(&self, name: &str) -> Option<&String> {
        get_prp(&self.prp, name)
    }

    /// Returns a property or a default.
    pub fn prp_def<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        get_prp_def(&self.prp, name, default)
    }

    /// Sets a textual content for this part. This is only used
    /// for text and currency-symbol.
    pub fn set_content<S: Into<String>>(&mut self, content: S) {
        self.content = Some(content.into());
    }

    /// Returns the text content.
    pub fn content(&self) -> Option<&String> {
        self.content.as_ref()
    }

    /// Tries to format the given boolean, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_boolean(&self, buf: &mut String, b: bool) {
        match self.part_type {
            FormatPartType::Boolean => {
                buf.push_str(if b { "true" } else { "false" });
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given float, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_float(&self, buf: &mut String, f: f64) {
        match self.part_type {
            FormatPartType::Number => {
                let dec = self.prp_def("number:decimal-places", "0").parse::<usize>();
                if let Ok(dec) = dec {
                    buf.push_str(&format!("{:.*}", dec, f));
                }
            }
            FormatPartType::Scientific => {
                buf.push_str(&format!("{:e}", f));
            }
            FormatPartType::CurrencySymbol => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given string, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_str(&self, buf: &mut String, s: &str) {
        match self.part_type {
            FormatPartType::TextContent => {
                buf.push_str(s);
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given DateTime, and appends the result to buf.
    /// Uses chrono::strftime for the implementation.
    /// If this part does'nt match does nothing
    #[allow(clippy::collapsible_if)]
    fn format_datetime(&self, buf: &mut String, d: &NaiveDateTime, h12: bool) {
        match self.part_type {
            FormatPartType::Day => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%d").to_string());
                } else {
                    buf.push_str(&d.format("%-d").to_string());
                }
            }
            FormatPartType::Month => {
                let is_long = self.prp_def("number:style", "") == "long";
                let is_text = self.prp_def("number:textual", "") == "true";
                if is_text {
                    if is_long {
                        buf.push_str(&d.format("%b").to_string());
                    } else {
                        buf.push_str(&d.format("%B").to_string());
                    }
                } else {
                    if is_long {
                        buf.push_str(&d.format("%m").to_string());
                    } else {
                        buf.push_str(&d.format("%-m").to_string());
                    }
                }
            }
            FormatPartType::Year => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%Y").to_string());
                } else {
                    buf.push_str(&d.format("%y").to_string());
                }
            }
            FormatPartType::DayOfWeek => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%A").to_string());
                } else {
                    buf.push_str(&d.format("%a").to_string());
                }
            }
            FormatPartType::WeekOfYear => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%W").to_string());
                } else {
                    buf.push_str(&d.format("%-W").to_string());
                }
            }
            FormatPartType::Hours => {
                let is_long = self.prp_def("number:style", "") == "long";
                if !h12 {
                    if is_long {
                        buf.push_str(&d.format("%H").to_string());
                    } else {
                        buf.push_str(&d.format("%-H").to_string());
                    }
                } else {
                    if is_long {
                        buf.push_str(&d.format("%I").to_string());
                    } else {
                        buf.push_str(&d.format("%-I").to_string());
                    }
                }
            }
            FormatPartType::Minutes => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%M").to_string());
                } else {
                    buf.push_str(&d.format("%-M").to_string());
                }
            }
            FormatPartType::Seconds => {
                let is_long = self.prp_def("number:style", "") == "long";
                if is_long {
                    buf.push_str(&d.format("%S").to_string());
                } else {
                    buf.push_str(&d.format("%-S").to_string());
                }
            }
            FormatPartType::AmPm => {
                buf.push_str(&d.format("%p").to_string());
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }

    /// Tries to format the given Duration, and appends the result to buf.
    /// If this part does'nt match does nothing
    fn format_time_duration(&self, buf: &mut String, d: &Duration) {
        match self.part_type {
            FormatPartType::Hours => {
                buf.push_str(&d.num_hours().to_string());
            }
            FormatPartType::Minutes => {
                buf.push_str(&(d.num_minutes() % 60).to_string());
            }
            FormatPartType::Seconds => {
                buf.push_str(&(d.num_seconds() % 60).to_string());
            }
            FormatPartType::Text => {
                if let Some(content) = &self.content {
                    buf.push_str(content)
                }
            }
            _ => {}
        }
    }
}

/// Creates a new number format.
pub fn create_boolean_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::Boolean);

    v.push_part(FormatPart::new(FormatPartType::Boolean));

    v
}

/// Creates a new number format.
pub fn create_number_format<S: Into<String>>(name: S, decimal: u8, grouping: bool) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::Number);

    let mut p = FormatPart::new(FormatPartType::Number);
    p.set_prp("number:min-integer-digits", 1.to_string());
    p.set_prp("number:decimal-places", decimal.to_string());
    p.set_prp("loext:min-decimal-places", 0.to_string());
    if grouping {
        p.set_prp("number:grouping", String::from("true"));
    }

    v.push_part(p);

    v
}

/// Creates a new number format with a fixed number of decimal places.
pub fn create_number_format_fixed<S: Into<String>>(name: S, decimal: u8, grouping: bool) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::Number);

    let mut p = FormatPart::new(FormatPartType::Number);
    p.set_prp("number:min-integer-digits", 1.to_string());
    p.set_prp("number:decimal-places", decimal.to_string());
    p.set_prp("loext:min-decimal-places", decimal.to_string());
    if grouping {
        p.set_prp("number:grouping", String::from("true"));
    }

    v.push_part(p);

    v
}

/// Creates a new percantage format.<
pub fn create_percentage_format<S: Into<String>>(name: S, decimal: u8) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::Percentage);

    let mut p = FormatPart::new(FormatPartType::Number);
    p.set_prp("number:min-integer-digits", 1.to_string());
    p.set_prp("number:decimal-places", decimal.to_string());
    p.set_prp("loext:min-decimal-places", decimal.to_string());
    v.push_part(p);

    let mut p2 = FormatPart::new(FormatPartType::Text);
    p2.set_content("&#160;%");
    v.push_part(p2);

    v
}

/// Creates a new currency format for EURO.
pub fn create_euro_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::Currency);

    let mut p0 = FormatPart::new(FormatPartType::CurrencySymbol);
    p0.set_prp("number:language", String::from("de"));
    p0.set_prp("number:country", String::from("AT"));
    p0.set_content("€");
    v.push_part(p0);

    let mut p1 = FormatPart::new(FormatPartType::Text);
    p1.set_content(" ");
    v.push_part(p1);

    let mut p2 = FormatPart::new(FormatPartType::Number);
    p2.set_prp("number:min-integer-digits", 1.to_string());
    p2.set_prp("number:decimal-places", 2.to_string());
    p2.set_prp("loext:min-decimal-places", 2.to_string());
    p2.set_prp("number:grouping", String::from("true"));
    v.push_part(p2);

    v
}

/// Creates a new currency format for EURO with negative values in red.
/// Needs the name of the positive format.
pub fn create_euro_red_format<S: Into<String>>(name: S, positive_style: S) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::Currency);

    let mut p0 = FormatPart::new(FormatPartType::StyleText);
    p0.set_prp("fo:color", String::from("#ff0000"));
    v.push_part(p0);

    let mut p1 = FormatPart::new(FormatPartType::Text);
    p1.set_content("-");
    v.push_part(p1);

    let mut p2 = FormatPart::new(FormatPartType::CurrencySymbol);
    p2.set_prp("number:language", String::from("de"));
    p2.set_prp("number:country", String::from("AT"));
    p2.set_content("€");
    v.push_part(p2);

    let mut p3 = FormatPart::new(FormatPartType::Text);
    p3.set_content(" ");
    v.push_part(p3);

    let mut p4 = FormatPart::new(FormatPartType::Number);
    p4.set_prp("number:min-integer-digits", 1.to_string());
    p4.set_prp("number:decimal-places", 2.to_string());
    p4.set_prp("loext:min-decimal-places", 2.to_string());
    p4.set_prp("number:grouping", String::from("true"));
    v.push_part(p4);

    let mut p5 = FormatPart::new(FormatPartType::StyleMap);
    p5.set_prp("style:condition", String::from("value()&gt;=0"));
    p5.set_prp("style:apply-style-name", positive_style.into());
    v.push_part(p5);

    v
}

/// Creates a new date format D.M.Y
pub fn create_date_dmy_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::DateTime);

    v.push_parts(vec![
        FormatPart::new_vec(FormatPartType::Day, vec![("number:style", String::from("long"))]),
        FormatPart::new_content(FormatPartType::Text, "."),
        FormatPart::new_vec(FormatPartType::Month, vec![("number:style", String::from("long"))]),
        FormatPart::new_content(FormatPartType::Text, "."),
        FormatPart::new_vec(FormatPartType::Year, vec![("number:style", String::from("long"))]),
    ]);

    v
}

/// Creates a datetime format Y-M-D H:M:S
pub fn create_datetime_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::DateTime);

    v.push_parts(vec![
        FormatPart::new_vec(FormatPartType::Year, vec![("number:style", String::from("long"))]),
        FormatPart::new_content(FormatPartType::Text, "-"),
        FormatPart::new_vec(FormatPartType::Month, vec![("number:style", String::from("long"))]),
        FormatPart::new_content(FormatPartType::Text, "-"),
        FormatPart::new_vec(FormatPartType::Day, vec![("number:style", String::from("long"))]),
        FormatPart::new_content(FormatPartType::Text, " "),
        FormatPart::new(FormatPartType::Hours),
        FormatPart::new_content(FormatPartType::Text, ":"),
        FormatPart::new(FormatPartType::Minutes),
        FormatPart::new_content(FormatPartType::Text, ":"),
        FormatPart::new(FormatPartType::Seconds),
    ]);

    v
}

/// Creates a new time-Duration format H:M:S
pub fn create_time_format<S: Into<String>>(name: S) -> ValueFormat {
    let mut v = ValueFormat::with_name(name.into(), ValueType::TimeDuration);

    v.push_parts(vec![
        FormatPart::new(FormatPartType::Hours),
        FormatPart::new_content(FormatPartType::Text, " "),
        FormatPart::new(FormatPartType::Minutes),
        FormatPart::new_content(FormatPartType::Text, " "),
        FormatPart::new(FormatPartType::Seconds),
    ]);

    v
}
