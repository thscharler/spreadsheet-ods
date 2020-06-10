use std::collections::{hash_map, HashMap};
use std::fmt::{Display, Formatter};

use color::Rgb;
use string_cache::DefaultAtom;

pub type AttrMapType = HashMap<DefaultAtom, String>;

pub trait AttrMap {
    /// Reference to the map of actual attributes.
    fn attr_map(&self) -> Option<&AttrMapType>;
    /// Reference to the map of actual attributes.
    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType>;

    /// Are there any attributes?
    fn is_empty(&self) -> bool {
        self.attr_map().is_none()
    }

    /// Add from Vec
    fn add_all(&mut self, data: Vec<(&str, String)>) {
        let attr = self.attr_map_mut();
        if attr.is_none() {
            attr.replace(HashMap::new());
        }
        if let Some(ref mut prp) = attr {
            for (name, val) in data {
                prp.insert(DefaultAtom::from(name), val);
            }
        }
    }

    /// Adds an attribute.
    fn set_attr(&mut self, name: &str, value: String) {
        let attr = self.attr_map_mut();
        if attr.is_none() {
            attr.replace(HashMap::new());
        }
        if let Some(ref mut attr) = attr {
            attr.insert(DefaultAtom::from(name), value);
        }
    }

    fn clear_attr(&mut self, name: &str) -> Option<String> {
        if let Some(ref mut attr) = self.attr_map_mut() {
            attr.remove(&DefaultAtom::from(name))
        } else {
            None
        }
    }

    fn attr(&self, name: &str) -> Option<&String> {
        if let Some(prp) = self.attr_map() {
            prp.get(&DefaultAtom::from(name))
        } else {
            None
        }
    }
}

pub struct AttrMapIter<'a> {
    it: Option<hash_map::Iter<'a, DefaultAtom, String>>,
}

impl<'a> AttrMapIter<'a> {
    pub fn from(attrmap: Option<&'a AttrMapType>) -> AttrMapIter<'a> {
        if let Some(attrmap) = attrmap {
            Self {
                it: Some(attrmap.into_iter()),
            }
        } else {
            Self {
                it: None,
            }
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

pub trait AttrFoBackground
    where Self: AttrMap {
    /// Border style.
    fn set_background_color(&mut self, color: Rgb<u8>) {
        self.set_attr("fo:background-color", color_string(color));
    }
}

pub trait AttrFoMinHeight
    where Self: AttrMap {
    fn set_min_height(&mut self, height: &str) {
        self.set_attr("fo:min-height", height.to_string());
    }
}

pub trait AttrFoMargin
    where Self: AttrMap {
    fn set_margin(&mut self, margin: &str) {
        self.set_attr("fo:margin", margin.to_string());
    }

    fn set_margin_bottom(&mut self, margin: &str) {
        self.set_attr("fo:margin-bottom", margin.to_string());
    }

    fn set_margin_left(&mut self, margin: &str) {
        self.set_attr("fo:margin-left", margin.to_string());
    }

    fn set_margin_right(&mut self, margin: &str) {
        self.set_attr("fo:margin-right", margin.to_string());
    }

    fn set_margin_top(&mut self, margin: &str) {
        self.set_attr("fo:margin-top", margin.to_string());
    }
}

pub trait AttrFoBorder
    where Self: AttrMap {
    /// Border style all four sides.
    fn border(&mut self, pt_width: f32, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border", border_string(pt_width, border, color));
    }

    /// Border style.
    fn border_bottom(&mut self, pt_width: f32, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-bottom", border_string(pt_width, border, color));
    }

    /// Border style.
    fn border_top(&mut self, pt_width: f32, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-top", border_string(pt_width, border, color));
    }

    /// Border style.
    fn border_left(&mut self, pt_width: f32, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-left", border_string(pt_width, border, color));
    }

    /// Border style.
    fn border_right(&mut self, pt_width: f32, border: Border, color: Rgb<u8>) {
        self.set_attr("fo:border-right", border_string(pt_width, border, color));
    }
}

pub trait AttrFoPadding
    where Self: AttrMap {
    fn set_padding(&mut self, padding: &str) {
        self.set_attr("fo:padding", padding.to_string());
    }

    fn set_padding_bottom(&mut self, padding: &str) {
        self.set_attr("fo:padding-bottom", padding.to_string());
    }

    fn set_padding_left(&mut self, padding: &str) {
        self.set_attr("fo:padding-left", padding.to_string());
    }

    fn set_padding_right(&mut self, padding: &str) {
        self.set_attr("fo:padding-right", padding.to_string());
    }

    fn set_padding_top(&mut self, padding: &str) {
        self.set_attr("fo:padding-top", padding.to_string());
    }
}

// format as string
#[allow(dead_code)]
fn color_string(color: Rgb<u8>) -> String {
    format!(" #{:02x}{:02x}{:02x}", color.r, color.g, color.b)
}

// format as string
#[allow(dead_code)]
fn border_string(width: f32, border: Border, color: Rgb<u8>) -> String {
    format!("{}pt {} #{:02x}{:02x}{:02x}", width, border, color.r, color.g, color.b)
}


/// Various border styles.
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