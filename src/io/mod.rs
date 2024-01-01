use crate::HashMap;
use loupe::{MemoryUsage, MemoryUsageTracker};
use std::borrow::Cow;

pub(crate) mod format;
pub(crate) mod parse;
pub(crate) mod read;
pub(crate) mod write;

mod xmlwriter;

const DUMP_XML: bool = false;
const DUMP_UNUSED: bool = false;

#[derive(Clone, Debug)]
pub(crate) struct NamespaceMap {
    map: HashMap<Cow<'static, str>, Cow<'static, str>>,
}

impl MemoryUsage for NamespaceMap {
    fn size_of_val(&self, _tracker: &mut dyn MemoryUsageTracker) -> usize {
        0
    }
}

impl NamespaceMap {
    pub(crate) fn new() -> Self {
        Self {
            map: Default::default(),
        }
    }

    pub(crate) fn insert(&mut self, k: String, v: String) {
        self.map.insert(Cow::Owned(k), Cow::Owned(v));
    }

    pub(crate) fn insert_str(&mut self, k: &'static str, v: &'static str) {
        self.map.insert(Cow::Borrowed(k), Cow::Borrowed(v));
    }

    pub(crate) fn entries(&self) -> impl Iterator<Item = (&Cow<'static, str>, &Cow<'static, str>)> {
        self.map.iter()
    }
}
