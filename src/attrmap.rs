//!
//! Defines the type AttrMap as container for different attribute-sets.
//! And there are a number of traits working with AttrMap to set
//! related families of attributes.
//!

use crate::sealed;
use std::collections::{hash_map, HashMap};
use string_cache::DefaultAtom;

/// Container type for attributes.
pub type AttrMapType = Option<Box<HashMap<DefaultAtom, String>>>;

/// Container trait for attributes.
///
/// Sealed
///
/// This trait is only to be used internally.
pub trait AttrMap: sealed::Sealed {
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
