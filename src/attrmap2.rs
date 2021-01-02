//!
//! Defines the type AttrMap as container for different attribute-sets.
//! And there are a number of traits working with AttrMap to set
//! related families of attributes.
//!

use std::collections::{hash_map, HashMap};
use string_cache::DefaultAtom;

type MapType = Option<Box<HashMap<DefaultAtom, String>>>;

/// Allows forwarding for structs that contain an AttrMap2
pub trait AttrMap2Trait {
    /// Reference to the map of actual attributes.
    fn attr_map(&self) -> &AttrMap2;
    /// Reference to the map of actual attributes.
    fn attr_map_mut(&mut self) -> &mut AttrMap2;

    /// Are there any attributes?
    fn is_empty(&self) -> bool {
        self.attr_map().is_empty()
    }

    /// Add from Vec
    fn add_all(&mut self, data: Vec<(&str, String)>) {
        self.attr_map_mut().add_all(data);
    }

    /// Adds an attribute.
    fn set_attr(&mut self, name: &str, value: String) {
        self.attr_map_mut().set_attr(name, value);
    }

    /// Removes an attribute.
    fn clear_attr(&mut self, name: &str) -> Option<String> {
        self.attr_map_mut().clear_attr(name)
    }

    /// Returns the attribute.
    fn attr(&self, name: &str) -> Option<&String> {
        self.attr_map().attr(name)
    }
}

/// Container type for attributes.
#[derive(Default, Clone, Debug)]
pub struct AttrMap2 {
    map: MapType,
}

impl AttrMap2 {
    pub fn new() -> Self {
        AttrMap2 {
            map: Default::default(),
        }
    }
    /// Are there any attributes?
    pub fn is_empty(&self) -> bool {
        self.map.is_none()
    }

    /// Add from Vec
    pub fn add_all(&mut self, data: Vec<(&str, String)>) {
        let attr = self.map.get_or_insert_with(|| Box::new(HashMap::new()));
        for (name, val) in data {
            attr.insert(DefaultAtom::from(name), val);
        }
    }

    /// Adds an attribute.
    pub fn set_attr(&mut self, name: &str, value: String) {
        self.map
            .get_or_insert_with(|| Box::new(HashMap::new()))
            .insert(DefaultAtom::from(name), value);
    }

    /// Removes an attribute.
    pub fn clear_attr(&mut self, name: &str) -> Option<String> {
        if let Some(ref mut attr) = self.map {
            attr.remove(&DefaultAtom::from(name))
        } else {
            None
        }
    }

    /// Returns the attribute.
    pub fn attr(&self, name: &str) -> Option<&String> {
        if let Some(ref prp) = self.map {
            prp.get(&DefaultAtom::from(name))
        } else {
            None
        }
    }

    pub fn iter(&self) -> AttrMapIter {
        From::from(self)
    }
}

/// Iterator for an AttrMap.
pub struct AttrMapIter<'a> {
    it: Option<hash_map::Iter<'a, DefaultAtom, String>>,
}

impl<'a> From<&'a AttrMap2> for AttrMapIter<'a> {
    fn from(attrmap: &'a AttrMap2) -> Self {
        if let Some(ref attrmap) = attrmap.map {
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
