use std::cell::RefCell;
use std::cmp::Ordering;
use std::collections::{BTreeMap, HashSet};
use std::fmt::{Debug, Display, Formatter};
use std::rc::Rc;

use crate::{ucell, OdsError, Sheet, Value};

#[derive(Debug)]
pub enum MapError<K, V> {
    InsertDuplicate(V, String),
    UniqueKeyViolation(String),
    NotUpdated(V, String),
    KeyError(K, String),
    ValueError(V, String),
}

impl<K, V> From<MapError<K, V>> for OdsError
where
    K: Debug,
    V: Debug,
{
    fn from(e: MapError<K, V>) -> Self {
        OdsError::Ods(e.to_string())
    }
}

impl<K, V> Display for MapError<K, V>
where
    K: Debug,
    V: Debug,
{
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            MapError::InsertDuplicate(v, msg) => {
                write!(f, "Duplicate key inserted: {} {:?}", msg, v)
            }
            MapError::NotUpdated(v, msg) => {
                write!(f, "Key already exists. Not inserted: {} {:?}", msg, v)
            }
            MapError::KeyError(k, msg) => {
                write!(f, "Key error: {} {:?}", msg, k)
            }
            MapError::ValueError(v, msg) => {
                write!(f, "Value error: {} {:?}", msg, v)
            }
            MapError::UniqueKeyViolation(msg) => {
                write!(f, "Unique key violation. {}", msg)
            }
        }
    }
}

/// Provides a basic view into the sheet data. Can shift the index for access.
#[derive(Debug)]
pub struct SheetView<'a> {
    sheet: &'a mut Sheet,
    drow: ucell,
    dcol: ucell,
}

#[allow(dead_code)]
impl<'a> SheetView<'a> {
    pub fn new(sheet: &'a mut Sheet, drow: ucell, dcol: ucell) -> Self {
        Self { sheet, drow, dcol }
    }

    /// Changes the value in the sheet.
    pub fn set_value<V: Into<Value>>(&mut self, row: ucell, col: ucell, value: V) {
        let row = self.drow + row;
        let col = self.dcol + col;
        self.sheet.set_value(row, col, value);
    }

    /// Gets the value from the sheet.
    pub fn value(&self, row: ucell, col: ucell) -> &Value {
        let row = self.drow + row;
        let col = self.dcol + col;
        self.sheet.value(row, col)
    }
}

/// Extracts further keys from the data. This is used by Index2 to
/// allow for extra indizes.
pub trait ExtractKey<K, V> {
    fn key<'a>(&self, val: &'a V) -> &'a K;
}

/// Any struct can implement this to load/store data from a row
/// in a sheet.
pub trait Recorder<K, V> {
    fn key<'a>(&self, val: &'a V) -> &'a K;

    /// Returns a header that is used for the sheet.
    fn def_header(&self) -> Option<&'static [&'static str]>;

    /// Loads from the sheet. None indicates there is no more data.
    fn load(&self, sheet: &SheetView, row: u32) -> Result<Option<V>, MapError<K, V>>;

    /// Stores to the sheet.
    fn store(&self, sheet: &mut SheetView, row: u32, val: &V) -> Result<(), MapError<K, V>>;
}

impl<T, K, V> ExtractKey<K, V> for T
where
    T: Recorder<K, V>,
{
    fn key<'a>(&self, val: &'a V) -> &'a K {
        Recorder::key(self, val)
    }
}

#[derive(Debug)]
pub enum IndexChecks {
    Fine,
    UniqueViolation,
    NotFound,
}

/// Links the extra indexes to the main storage.
pub trait IndexBackend<V> {
    fn name(&self) -> &str;

    /// Clears the index.
    fn clear(&mut self);

    /// Checks for any constraint violations if we would insert this.
    #[must_use]
    fn check(&mut self, value: &V, idx: usize) -> IndexChecks;

    /// A value has been inserted.
    fn insert(&mut self, value: &V, idx: usize);

    /// A value has been removed.
    fn remove(&mut self, value: &V, idx: usize);
}

/// Implements an extra index into the data.
pub struct Index1<K, V>
where
    K: Ord + Clone,
{
    name: String,
    extract_key: Box<dyn ExtractKey<K, V> + 'static>,
    index: BTreeMap<K, usize>,
}

impl<K, V> Debug for Index1<K, V>
where
    K: Ord + Clone + Debug,
{
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "Index1 {}", self.name)?;
        writeln!(f, "{:?}", self.index)?;
        Ok(())
    }
}

impl<K1, V> Index1<K1, V>
where
    K1: Ord + Clone,
{
    /// Creates an extra index for the data.
    /// The index must be added to the matching MapSheet to be active.
    pub fn new<I: 'static + ExtractKey<K1, V>>(extract: I) -> Rc<RefCell<Index1<K1, V>>> {
        Rc::new(RefCell::new(Self {
            name: "".to_string(),
            extract_key: Box::new(extract),
            index: Default::default(),
        }))
    }

    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_string()
    }

    /// Size.
    pub fn len(&self) -> usize {
        self.index.len()
    }

    /// Contains.
    pub fn contains_key(&self, key: &K1) -> bool {
        self.find_idx(key).is_some()
    }

    /// Returns the indexes where this key occurs.
    pub fn find_idx(&self, key: &K1) -> Option<usize> {
        if let Some(idx) = self.index.get(key) {
            Some(*idx)
        } else {
            None
        }
    }

    /// Iterate.
    pub fn iter(&self) -> std::collections::btree_map::Iter<K1, usize> {
        self.index.iter()
    }

    // todo: more
}

impl<K, V> IndexBackend<V> for Index1<K, V>
where
    K: Ord + Clone,
{
    fn name(&self) -> &str {
        self.name.as_ref()
    }

    /// Function for MapSheet to clear the index.
    fn clear(&mut self) {
        self.index.clear();
    }

    /// Checks for any constraint violations.
    fn check(&mut self, value: &V, _idx: usize) -> IndexChecks {
        let key = self.extract_key.key(value);
        if self.index.contains_key(key) {
            IndexChecks::UniqueViolation
        } else {
            IndexChecks::Fine
        }
    }

    /// Inserts a value to the index.
    fn insert(&mut self, value: &V, idx: usize) {
        let key = self.extract_key.key(value);
        if self.index.insert(key.clone(), idx).is_some() {
            panic!("cuplicate key inserted");
        }
    }

    /// Removes a value from the index.
    fn remove(&mut self, value: &V, _idx: usize) {
        let key = self.extract_key.key(value);
        self.index.remove(key);
    }
}

/// Implements an extra index into the data.
pub struct Index2<K, V>
where
    K: Ord + Clone,
{
    name: String,
    extract_key: Box<dyn ExtractKey<K, V> + 'static>,
    index: BTreeMap<K, HashSet<usize>>,
}

impl<K, V> Debug for Index2<K, V>
where
    K: Ord + Clone + Debug,
{
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "Index2 {}", self.name)?;
        writeln!(f, "{:?}", self.index)?;
        Ok(())
    }
}

impl<K2, V> Index2<K2, V>
where
    K2: Ord + Clone,
{
    /// Creates an extra index for the data.
    /// The index must be added to the matching MapSheet to be active.
    pub fn new<I: 'static + ExtractKey<K2, V>>(extract: I) -> Rc<RefCell<Index2<K2, V>>> {
        Rc::new(RefCell::new(Self {
            name: "".to_string(),
            extract_key: Box::new(extract),
            index: Default::default(),
        }))
    }

    pub fn set_name(&mut self, name: &str) {
        self.name = name.to_string();
    }

    /// Size.
    pub fn len(&self) -> usize {
        self.index.len()
    }

    /// Contains.
    pub fn contains_key(&self, key: &K2) -> bool {
        self.find_idx(key).is_some()
    }

    /// Returns the indexes where this key occurs.
    pub fn find_idx(&self, key: &K2) -> Option<&HashSet<usize>> {
        if let Some(idx) = self.index.get(key) {
            Some(idx)
        } else {
            None
        }
    }

    /// Iterate.
    pub fn iter(&self) -> std::collections::btree_map::Iter<K2, HashSet<usize>> {
        self.index.iter()
    }

    // todo: more
}

impl<K, V> IndexBackend<V> for Index2<K, V>
where
    K: Ord + Clone,
{
    fn name(&self) -> &str {
        &self.name
    }

    /// Function for MapSheet to clear the index.
    fn clear(&mut self) {
        self.index.clear();
    }

    /// Checks for any constraint violations.
    fn check(&mut self, _value: &V, _idx: usize) -> IndexChecks {
        IndexChecks::Fine
    }

    /// Inserts a value to the index.
    fn insert(&mut self, value: &V, idx: usize) {
        let key = self.extract_key.key(value);
        let row_set = self
            .index
            .entry(key.clone())
            .or_insert_with(|| HashSet::new());
        row_set.insert(idx);
    }

    /// Removes a value from the index.
    fn remove(&mut self, value: &V, idx: usize) {
        let key = self.extract_key.key(value);
        let row_set = self.index.get_mut(key);
        if let Some(row_set) = row_set {
            row_set.remove(&idx);
            if row_set.is_empty() {
                self.index.remove(key);
            }
        }
    }
}

/// Allows to access row oriented data in a sheet.
///
/// There is a mapping trait Recorder to load/store the data.
///
/// A primary index is always kept, and more indexes can be attached.
///
/// Constraints
/// The inserted value can be modified in place, but any value that is
/// part of a index must not be touched. For those cases a clone should be
/// made and the update function be called.
///
pub struct MapSheet<K, V>
where
    K: Ord + Clone,
{
    /// Mapping trait.
    recorder: Box<dyn Recorder<K, V> + 'static>,
    /// Data. Deletes only set the option to None to keep the indexes alive.
    data: Vec<Option<V>>,
    /// Primary index in the data.
    primary_index: Rc<RefCell<Index1<K, V>>>,
    /// Extra indexes.
    indexes: Vec<Rc<RefCell<dyn IndexBackend<V>>>>,
}

impl<K, V> Debug for MapSheet<K, V>
where
    K: Ord + Clone + Debug,
    V: Debug,
{
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "MapSheet (")?;
        writeln!(f, "    data {:?}", self.data)?;
        writeln!(f, "    primary {:?}", self.primary_index)?;
        writeln!(f, "    indexes {:?}", self.indexes.len())?;
        Ok(())
    }
}

#[allow(dead_code)]
impl<K, V> MapSheet<K, V>
where
    K: Ord + Clone,
{
    /// Creates a new map.
    pub fn new<R: Recorder<K, V> + Clone + 'static>(record: R) -> MapSheet<K, V> {
        Self {
            recorder: Box::new(record.clone()),
            primary_index: Index1::new(record),
            data: Default::default(),
            indexes: vec![],
        }
    }

    /// Links an index.
    pub fn add_index(
        &mut self,
        index: Rc<RefCell<dyn IndexBackend<V>>>,
    ) -> Result<(), MapError<K, V>> {
        // Insert into this index.
        for (idx, value) in self.data.iter().enumerate() {
            if let Some(value) = value {
                let mut index = index.as_ref().borrow_mut();
                match index.check(&value, idx) {
                    IndexChecks::Fine => {}
                    IndexChecks::UniqueViolation => {
                        return Err(MapError::UniqueKeyViolation(format!(
                            "Index {}",
                            index.name()
                        )));
                    }
                    IndexChecks::NotFound => {
                        unreachable!()
                    }
                }

                index.insert(value, idx);
            }
        }

        self.indexes.push(index);

        Ok(())
    }

    /// Returns the number of values in the map.
    pub fn len(&self) -> usize {
        self.primary_index.borrow().len()
    }

    /// Returns the length of the underlying data vec. Not the same as
    /// the number of actual values in this map.
    pub fn len_vec(&self) -> usize {
        self.data.len()
    }

    /// Access to the underlying data vec. May return None if the value
    /// was deleted.
    pub fn get_idx(&self, idx: usize) -> Option<&V> {
        if let Some(value) = self.data.get(idx) {
            value.as_ref()
        } else {
            None
        }
    }

    /// Access to the underlying data vec. May return None if the value
    /// was delete.
    pub fn get_idx_mut(&mut self, idx: usize) -> Option<&mut V> {
        if let Some(value) = self.data.get_mut(idx) {
            value.as_mut()
        } else {
            None
        }
    }

    /// Removes via the index.
    pub fn remove_idx(&mut self, idx: usize) -> Option<V> {
        if let Some(value) = self.data.get_mut(idx) {
            // changes the value to None
            if let Some(value) = value.take() {
                // clear indexes
                {
                    let mut pkindex = self.primary_index.as_ref().borrow_mut();
                    pkindex.remove(&value, idx);
                }
                for index in &self.indexes {
                    let mut index = index.as_ref().borrow_mut();
                    index.remove(&value, idx);
                }
                Some(value)
            } else {
                // already removed, fine
                None
            }
        } else {
            // index error, fine
            None
        }
    }

    /// Finds via the primary key.
    pub fn find_idx(&self, pkey: &K) -> Option<usize> {
        self.primary_index.borrow().find_idx(pkey)
    }

    /// Finds via the primary key.
    pub fn get_mut(&mut self, pkey: &K) -> Option<&mut V> {
        if let Some(idx) = self.primary_index.borrow().find_idx(pkey) {
            if let Some(value) = self.data.get_mut(idx) {
                value.as_mut()
            } else {
                panic!("key was already deleted, but could be found?");
            }
        } else {
            // index error, fine
            None
        }
    }

    /// Finds via the primary key.
    pub fn get(&self, pkey: &K) -> Option<&V> {
        if let Some(idx) = self.primary_index.borrow().find_idx(pkey) {
            if let Some(value) = self.data.get(idx) {
                value.as_ref()
            } else {
                panic!("key was already deleted, but could be found?");
            }
        } else {
            // index error, fine
            None
        }
    }

    /// Checks everything before inserting data.
    fn insertion_checks(&mut self, value: &V, idx: usize) -> bool {
        let mut pkindex = self.primary_index.as_ref().borrow_mut();
        match pkindex.check(&value, idx) {
            IndexChecks::Fine => {}
            IndexChecks::UniqueViolation => return false,
            _ => unreachable!(),
        }
        for index in &self.indexes {
            let mut index = index.as_ref().borrow_mut();
            match index.check(&value, idx) {
                IndexChecks::Fine => {}
                IndexChecks::UniqueViolation => return false,
                _ => unreachable!(),
            }
        }
        true
    }

    /// Checks everything before inserting data.
    fn insertion_fail(&mut self, value: V, idx: usize) -> Result<(), MapError<K, V>> {
        let mut pkindex = self.primary_index.as_ref().borrow_mut();
        match pkindex.check(&value, idx) {
            IndexChecks::Fine => {}
            IndexChecks::UniqueViolation => {
                return Err(MapError::InsertDuplicate(
                    value,
                    format!("Index constraint for {}", pkindex.name()),
                ))
            }
            _ => unreachable!(),
        }
        for index in &self.indexes {
            let mut index = index.as_ref().borrow_mut();
            match index.check(&value, idx) {
                IndexChecks::Fine => {}
                IndexChecks::UniqueViolation => {
                    return Err(MapError::InsertDuplicate(
                        value,
                        format!("Index constraint for {}", index.name()),
                    ))
                }
                _ => unreachable!(),
            }
        }
        Ok(())
    }

    /// Adds a new value.
    /// Fails if the primary key already exists.
    pub fn insert(&mut self, value: V) -> Result<(), MapError<K, V>> {
        let idx = self.data.len();

        // check
        if !self.insertion_checks(&value, idx) {
            return self.insertion_fail(value, idx);
        }

        // modify
        {
            let mut pkindex = self.primary_index.as_ref().borrow_mut();
            pkindex.insert(&value, idx);
        }
        for index in &self.indexes {
            let mut index = index.as_ref().borrow_mut();
            index.insert(&value, idx);
        }

        self.data.push(Some(value));

        Ok(())
    }

    /// Updates the record with the *old* key and modifies the value.
    /// Returns the old value on success.
    pub fn update(&mut self, old_key: &K, new_value: V) -> Result<Option<V>, MapError<K, V>> {
        // Does the old key exist?
        let idx = { self.primary_index.borrow().find_idx(old_key) };

        if let Some(idx) = idx {
            // check
            // the new key must not exist, if it changed.
            let new_key = self.recorder.key(&new_value);
            if old_key.cmp(new_key) != Ordering::Equal {
                if !self.insertion_checks(&new_value, idx) {
                    self.insertion_fail(new_value, idx)?;
                    unreachable!()
                }
            }

            // modify
            // extract the old value
            let old_value = self.data[idx]
                .take()
                .expect("key was already deleted, but could be found?");

            let mut pkindex = self.primary_index.as_ref().borrow_mut();
            // remove from indexes
            pkindex.remove(&old_value, idx);
            for index in &self.indexes {
                let mut index = index.as_ref().borrow_mut();
                index.remove(&old_value, idx);
            }

            // add to indexes
            pkindex.insert(&new_value, idx);
            for index in &self.indexes {
                let mut index = index.as_ref().borrow_mut();
                index.insert(&new_value, idx);
            }

            self.data[idx] = Some(new_value);

            Ok(Some(old_value))
        } else {
            Err(MapError::NotUpdated(new_value, String::new()))
        }
    }

    /// Removes a value via the primary key.
    pub fn remove(&mut self, key: &K) -> Option<V> {
        let idx = self.find_idx(key);
        if let Some(idx) = idx {
            self.remove_idx(idx)
        } else {
            None
        }
    }

    // todo: find, iter, ...

    /// Stores the data back to the sheet. The data is stored in the
    /// ordering of the primary key.
    pub fn store(&mut self, sheet: &mut Sheet) -> Result<(), MapError<K, V>> {
        // store header
        if let Some(header) = self.recorder.def_header() {
            for (col, hd) in header.iter().enumerate() {
                sheet.set_value(0, col as ucell, *hd);
            }
        }

        // header data
        let row = if self.recorder.def_header().is_none() {
            0
        } else {
            1
        };
        // create a view into the sheet. adjust for a header.
        let mut view = SheetView::new(sheet, row, 0);

        let mut row = 0u32;
        for (_key, idx) in self.primary_index.borrow().iter() {
            let val = self.data.get(*idx);
            // maybe not found
            if let Some(val) = val {
                // or removed
                if let Some(val) = val {
                    self.recorder.store(&mut view, row, val)?;
                }
            }
            row += 1;
        }

        // clean up any further data
        let mut keys = Vec::new();
        // todo: maybe impl for sheet?
        for (k, _) in sheet.data.range((row, 0)..) {
            keys.push(k.clone());
        }
        for k in keys {
            sheet.remove_cell(k.0, k.1);
        }

        Ok(())
    }

    /// Loads the data from the sheet.
    pub fn load(&mut self, sheet: &mut Sheet) -> Result<(), MapError<K, V>> {
        // reset
        self.data = Default::default();

        {
            let mut pkindex = self.primary_index.as_ref().borrow_mut();
            pkindex.clear();
        }
        for index in &self.indexes {
            let mut index = index.as_ref().borrow_mut();
            index.clear();
        }

        // header data
        let row = if self.recorder.def_header().is_none() {
            0
        } else {
            1
        };

        // create a view into the sheet. ignore the header, if any.
        let mut view = SheetView::new(sheet, row, 0);

        let mut row = 0;
        loop {
            let value = self.recorder.load(&mut view, row)?;

            if let Some(value) = value {
                // check
                if !self.insertion_checks(&value, row as usize) {
                    return self.insertion_fail(value, row as usize);
                }

                // modify
                {
                    let mut pkindex = self.primary_index.as_ref().borrow_mut();
                    pkindex.insert(&value, row as usize);
                }
                for index in &self.indexes {
                    let mut index = index.as_ref().borrow_mut();
                    index.insert(&value, row as usize);
                }

                self.data.push(Some(value));
            } else {
                break;
            }

            row += 1;
        }

        Ok(())
    }
}
