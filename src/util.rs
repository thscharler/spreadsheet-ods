use std::collections::HashMap;

use string_cache::DefaultAtom;

// copy the vector to a property-map.
pub(crate) fn set_prp_vec(map: &mut Option<HashMap<DefaultAtom, String>>, vec: Vec<(&str, String)>) {
    if map.is_none() {
        map.replace(HashMap::new());
    }
    if let Some(map) = map {
        for (name, value) in vec {
            let a = DefaultAtom::from(name);
            map.insert(a, value);
        }
    }
}

// set a property
pub(crate) fn set_prp(map: &mut Option<HashMap<DefaultAtom, String>>, name: &str, value: String) {
    if map.is_none() {
        map.replace(HashMap::new());
    }
    if let Some(map) = map {
        let a = DefaultAtom::from(name);
        map.insert(a, value);
    }
}

// return a property
pub(crate) fn get_prp<'a, 'b>(map: &'a Option<HashMap<DefaultAtom, String>>, name: &'b str) -> Option<&'a String> {
    if let Some(map) = map {
        map.get(&DefaultAtom::from(name))
    } else {
        None
    }
}

// return a property
pub(crate) fn get_prp_def<'a>(map: &'a Option<HashMap<DefaultAtom, String>>, name: &str, default: &'a str) -> &'a str {
    if let Some(map) = map {
        if let Some(value) = map.get(&DefaultAtom::from(name)) {
            value.as_ref()
        } else {
            default
        }
    } else {
        default
    }
}

