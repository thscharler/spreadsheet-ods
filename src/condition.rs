use std::fmt::{Display, Formatter};

use crate::CellRange;

//
// #[derive(Copy, Clone, Debug)]
// pub enum Cmp {
//     Equal,
//     NotEqual,
//     Less,
//     Greater,
//     LessOrEqual,
//     GreaterOrEqual,
// }
//
// impl Display for Cmp {
//     fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
//         match self {
//             Cmp::Equal => write!(f, "="),
//             Cmp::NotEqual => write!(f, "!="),
//             Cmp::Less => write!(f, "<"),
//             Cmp::Greater => write!(f, ">"),
//             Cmp::LessOrEqual => write!(f, "<="),
//             Cmp::GreaterOrEqual => write!(f, ">="),
//         }
//     }
// }

#[derive(Clone, Debug)]
pub struct Value {
    val: String,
}

fn quote(val: &str) -> String {
    let mut buf = String::new();
    buf.push('"');
    for c in val.chars() {
        if c == '"' {
            buf.push('"');
            buf.push('"');
        } else {
            buf.push(c);
        }
    }
    buf.push('"');
    buf
}

impl Display for Value {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.val)
    }
}

impl From<&str> for Value {
    fn from(s: &str) -> Self {
        Value { val: quote(s) }
    }
}

impl From<String> for Value {
    fn from(s: String) -> Self {
        Value {
            val: quote(s.as_str()),
        }
    }
}

impl From<&String> for Value {
    fn from(s: &String) -> Self {
        Value {
            val: quote(s.as_str()),
        }
    }
}

macro_rules! from_x_conditionvalue {
    ($int:ty) => {
        impl From<$int> for Value {
            fn from(v: $int) -> Self {
                Value { val: v.to_string() }
            }
        }

        impl From<&$int> for Value {
            fn from(v: &$int) -> Self {
                Value { val: v.to_string() }
            }
        }
    };
}

from_x_conditionvalue!(i8);
from_x_conditionvalue!(i16);
from_x_conditionvalue!(i32);
from_x_conditionvalue!(i64);
from_x_conditionvalue!(u8);
from_x_conditionvalue!(u16);
from_x_conditionvalue!(u32);
from_x_conditionvalue!(u64);
from_x_conditionvalue!(f32);
from_x_conditionvalue!(f64);
from_x_conditionvalue!(bool);

///
/// Conditions for Validations and StyleMaps.
///
#[derive(Clone, Debug)]
pub struct ValueCondition {
    cond: String,
}

impl Display for ValueCondition {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.cond)
    }
}

impl ValueCondition {
    pub fn new<S: Into<String>>(str: S) -> Self {
        Self { cond: str.into() }
    }

    /// Refers to the content.
    pub fn content_eq<V: Into<Value>>(value: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content()=");
        buf.push_str(value.into().to_string().as_str());
        ValueCondition { cond: buf }
    }

    /// Refers to the content.
    pub fn content_ne<V: Into<Value>>(value: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content()!=");
        buf.push_str(value.into().to_string().as_str());
        ValueCondition { cond: buf }
    }

    /// Refers to the content.
    pub fn content_lt<V: Into<Value>>(value: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content()<");
        buf.push_str(value.into().to_string().as_str());
        ValueCondition { cond: buf }
    }

    /// Refers to the content.
    pub fn content_gt<V: Into<Value>>(value: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content()>");
        buf.push_str(value.into().to_string().as_str());
        ValueCondition { cond: buf }
    }

    /// Refers to the content.
    pub fn content_lte<V: Into<Value>>(value: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content()<=");
        buf.push_str(value.into().to_string().as_str());
        ValueCondition { cond: buf }
    }

    /// Refers to the content.
    pub fn content_gte<V: Into<Value>>(value: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content()>=");
        buf.push_str(value.into().to_string().as_str());
        ValueCondition { cond: buf }
    }

    /// Range check.
    pub fn content_is_between<V: Into<Value>>(from: V, to: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-between(");
        buf.push_str(from.into().to_string().as_str());
        buf.push_str(", ");
        buf.push_str(to.into().to_string().as_str());
        buf.push_str(")");
        ValueCondition { cond: buf }
    }

    /// Range check.
    pub fn content_is_not_between<V: Into<Value>>(from: V, to: V) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-not-between(");
        buf.push_str(from.into().to_string().as_str());
        buf.push_str(", ");
        buf.push_str(to.into().to_string().as_str());
        buf.push_str(")");
        ValueCondition { cond: buf }
    }

    /// Evaluates a formula.
    pub fn is_true_formula<S: AsRef<str>>(formula: S) -> ValueCondition {
        let mut buf = String::new();
        buf.push_str("is-true-formula(");
        buf.push_str(formula.as_ref());
        buf.push_str(")");
        ValueCondition { cond: buf }
    }
}

///
/// Conditions for Validations and StyleMaps.
///
#[derive(Clone, Debug)]
pub struct Condition {
    cond: String,
}

impl Display for Condition {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.cond)
    }
}

impl Condition {
    pub fn new<S: Into<String>>(str: S) -> Self {
        Self { cond: str.into() }
    }

    /// Content length.
    pub fn content_text_length_eq<V: Into<Value>>(value: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length()=");
        buf.push_str(value.into().to_string().as_str());
        Condition { cond: buf }
    }

    /// Content length.
    pub fn content_text_length_ne<V: Into<Value>>(value: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length()!=");
        buf.push_str(value.into().to_string().as_str());
        Condition { cond: buf }
    }

    /// Content length.
    pub fn content_text_length_lt<V: Into<Value>>(value: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length()<");
        buf.push_str(value.into().to_string().as_str());
        Condition { cond: buf }
    }

    /// Content length.
    pub fn content_text_length_gt<V: Into<Value>>(value: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length()>");
        buf.push_str(value.into().to_string().as_str());
        Condition { cond: buf }
    }

    /// Content length.
    pub fn content_text_length_lte<V: Into<Value>>(value: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length()<=");
        buf.push_str(value.into().to_string().as_str());
        Condition { cond: buf }
    }

    /// Content length.
    pub fn content_text_length_gte<V: Into<Value>>(value: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length()>=");
        buf.push_str(value.into().to_string().as_str());
        Condition { cond: buf }
    }

    /// Range check.
    pub fn content_text_length_is_between<V: Into<Value>>(from: V, to: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length-is-between(");
        buf.push_str(from.into().to_string().as_str());
        buf.push_str(", ");
        buf.push_str(to.into().to_string().as_str());
        buf.push_str(")");
        Condition { cond: buf }
    }

    /// Range check.
    pub fn content_text_length_is_not_between<V: Into<Value>>(from: V, to: V) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-text-length-is-not-between(");
        buf.push_str(from.into().to_string().as_str());
        buf.push_str(", ");
        buf.push_str(to.into().to_string().as_str());
        buf.push_str(")");
        Condition { cond: buf }
    }

    /// The value is in this list.
    pub fn content_is_in_list<'a, V>(list: &'a [V]) -> Condition
    where
        Value: From<&'a V>,
    {
        let mut buf = String::new();
        buf.push_str("cell-content-is-in-list(");

        let mut sep = false;
        for v in list {
            if sep {
                buf.push_str(";");
            }
            buf.push('"');
            let vv: Value = v.into();
            buf.push_str(vv.to_string().as_str());
            buf.push('"');
            sep = true;
        }

        buf.push(')');
        Condition { cond: buf }
    }

    /// The choices are made up from the values in the cellrange.
    /// Warning
    /// For the cellrange the distance to the base-cell is calculated,
    /// and this result is added to the cell this condition is applied to.
    ///
    pub fn content_is_in_cellrange(range: CellRange) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-in-list(");
        buf.push_str(range.to_formula().as_str());
        buf.push(')');
        Condition { cond: buf }
    }

    /// Content is a date.
    pub fn content_is_date_and(vcond: ValueCondition) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-date()");
        buf.push_str(" and ");
        buf.push_str(vcond.to_string().as_str());
        Condition { cond: buf }
    }

    /// Content is a time.
    pub fn content_is_time_and(vcond: ValueCondition) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-time()");
        buf.push_str(" and ");
        buf.push_str(vcond.to_string().as_str());
        Condition { cond: buf }
    }

    /// Content is a number.
    pub fn content_is_decimal_number_and(vcond: ValueCondition) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-decimal-number()");
        buf.push_str(" and ");
        buf.push_str(vcond.to_string().as_str());
        Condition { cond: buf }
    }

    /// Content is a whole number.
    pub fn content_is_whole_number_and(vcond: ValueCondition) -> Condition {
        let mut buf = String::new();
        buf.push_str("cell-content-is-whole-number()");
        buf.push_str(" and ");
        buf.push_str(vcond.to_string().as_str());
        Condition { cond: buf }
    }

    /// Evaluates a formula.
    pub fn is_true_formula<S: AsRef<str>>(formula: S) -> Condition {
        let mut buf = String::new();
        buf.push_str("is-true-formula(");
        buf.push_str(formula.as_ref());
        buf.push_str(")");
        Condition { cond: buf }
    }
}

#[cfg(test)]
mod tests {
    use crate::condition::{Condition, ValueCondition};
    use crate::CellRange;

    #[test]
    pub fn test_valuecondition() {
        let c = ValueCondition::content_eq(5);
        assert_eq!(c.to_string(), "cell-content()=5");
        let c = ValueCondition::content_ne(5);
        assert_eq!(c.to_string(), "cell-content()!=5");
        let c = ValueCondition::content_lt(5);
        assert_eq!(c.to_string(), "cell-content()<5");
        let c = ValueCondition::content_gt(5);
        assert_eq!(c.to_string(), "cell-content()>5");
        let c = ValueCondition::content_lte(5);
        assert_eq!(c.to_string(), "cell-content()<=5");
        let c = ValueCondition::content_gte(5);
        assert_eq!(c.to_string(), "cell-content()>=5");
        let c = ValueCondition::content_is_not_between(1, 5);
        assert_eq!(c.to_string(), "cell-content-is-not-between(1, 5)");
        let c = ValueCondition::content_is_between(1, 5);
        assert_eq!(c.to_string(), "cell-content-is-between(1, 5)");
        let c = ValueCondition::is_true_formula("formula");
        assert_eq!(c.to_string(), "is-true-formula(formula)");
    }

    #[test]
    fn test_condition() {
        let c = Condition::content_text_length_eq(7);
        assert_eq!(c.to_string(), "cell-content-text-length()=7");
        let c = Condition::content_text_length_is_between(5, 7);
        assert_eq!(c.to_string(), "cell-content-text-length-is-between(5, 7)");
        let c = Condition::content_text_length_is_not_between(5, 7);
        assert_eq!(
            c.to_string(),
            "cell-content-text-length-is-not-between(5, 7)"
        );
        let c = Condition::content_is_in_list(&[1, 2, 3, 4, 5]);
        assert_eq!(
            c.to_string(),
            r#"cell-content-is-in-list("1";"2";"3";"4";"5")"#
        );
        let c = Condition::content_is_in_cellrange(CellRange::remote("other", 0, 0, 10, 0));
        assert_eq!(c.to_string(), "cell-content-is-in-list([other.A1:.A11])");

        let c = Condition::content_is_date_and(ValueCondition::content_eq(0));
        assert_eq!(c.to_string(), "cell-content-is-date() and cell-content()=0");
        let c = Condition::content_is_time_and(ValueCondition::content_eq(0));
        assert_eq!(c.to_string(), "cell-content-is-time() and cell-content()=0");
        let c = Condition::content_is_decimal_number_and(ValueCondition::content_eq(0));
        assert_eq!(
            c.to_string(),
            "cell-content-is-decimal-number() and cell-content()=0"
        );
        let c = Condition::content_is_whole_number_and(ValueCondition::content_eq(0));
        assert_eq!(
            c.to_string(),
            "cell-content-is-whole-number() and cell-content()=0"
        );

        let c = Condition::is_true_formula("formula");
        assert_eq!(c.to_string(), "is-true-formula(formula)");
    }
}
