use crate::{FormatPart, FormatPartType, ValueFormat, ValueType};

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