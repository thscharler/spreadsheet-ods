use std::cell::RefCell;
use std::collections::BTreeMap;
use std::fmt::{Debug, Formatter};
use std::marker::PhantomData;

use crate::{ucell, Sheet, Value};

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

    pub fn set_value<V: Into<Value>>(&mut self, row: ucell, col: ucell, value: V) {
        let row = self.drow + row;
        let col = self.dcol + col;
        self.sheet.set_value(row, col, value);
    }

    pub fn value(&self, row: ucell, col: ucell) -> &Value {
        let row = self.drow + row;
        let col = self.dcol + col;
        self.sheet.value(row, col)
    }
}

pub trait Record<K, V> {
    fn primary(&self, val: &V) -> K;
    fn def_header(&self) -> Option<&'static [&'static str]>;
    fn load(&self, sheet: &SheetView, n: u32) -> Option<V>;
    fn store(&self, sheet: &mut SheetView, n: u32, val: &V);
}

pub trait IndexValueToKey<V, K>
where
    K: Clone,
{
    fn key<'a>(&self, val: &'a V) -> &'a K;
}

pub trait IndexBackend<V> {
    fn clear(&mut self);
    fn insert(&mut self, value: &V, idx: usize);
    fn remove(&mut self, value: &V);
}

pub struct Index2<I, K, V>
where
    I: IndexValueToKey<V, K> + Debug,
    K: Ord + Clone + Debug,
    V: Debug,
{
    indexer: I,
    index: BTreeMap<K, usize>,
    data: PhantomData<V>,
}

impl<I, K, V> Debug for Index2<I, K, V>
where
    I: IndexValueToKey<V, K> + Debug,
    K: Ord + Clone + Debug,
    V: Debug,
{
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "Index2 (")?;
        writeln!(f, "{:?}", self.index)?;
        Ok(())
    }
}

impl<I, K2, V> Index2<I, K2, V>
where
    I: IndexValueToKey<V, K2> + Debug,
    K2: Ord + Clone + Debug,
    V: Debug,
{
    pub fn new(indexer: I) -> RefCell<Index2<I, K2, V>> {
        RefCell::new(Self {
            indexer,
            index: Default::default(),
            data: Default::default(),
        })
    }

    pub fn find(&self, key: K2) -> Option<usize> {
        if let Some(idx) = self.index.get(&key) {
            Some(*idx)
        } else {
            None
        }
    }

    fn clear_impl(&mut self) {
        self.index.clear();
    }

    fn insert_impl(&mut self, value: &V, idx: usize) {
        let key = self.indexer.key(value);
        self.index.insert(key.clone(), idx);
    }

    fn remove_impl(&mut self, value: &V) {
        let key = self.indexer.key(value);
        self.index.remove(&key);
    }
}

impl<I, K, V> IndexBackend<V> for Index2<I, K, V>
where
    I: IndexValueToKey<V, K> + Debug,
    K: Ord + Clone + Debug,
    V: Debug,
{
    fn clear(&mut self) {
        self.clear_impl();
    }

    fn insert(&mut self, value: &V, idx: usize) {
        self.insert_impl(value, idx);
    }

    fn remove(&mut self, value: &V) {
        self.remove_impl(value);
    }
}

pub fn link_index<'a, I, K1, R, K2, V>(
    index: &'a RefCell<Index2<I, K1, V>>,
    mapsheet: &mut MapSheet<'a, R, K2, V>,
) where
    I: IndexValueToKey<V, K1> + Debug + 'static,
    K1: Ord + Clone + Debug + 'static,
    R: Record<K2, V> + Debug,
    K2: Ord + Clone + Debug,
    V: Debug + 'static,
{
    let other_self: &RefCell<dyn IndexBackend<V>> = index;
    mapsheet.add_index(other_self);
}

pub struct MapSheet<'a, R, K, V>
where
    R: Record<K, V> + Debug,
    K: Ord + Clone + Debug,
    V: Debug,
{
    record: R,
    load_size: usize,
    data: Vec<Option<V>>,
    primary_index: BTreeMap<K, usize>,
    indexes: Vec<&'a RefCell<dyn IndexBackend<V>>>,
}

impl<'a, R, K, V> Debug for MapSheet<'a, R, K, V>
where
    R: Record<K, V> + Debug,
    K: Ord + Clone + Debug,
    V: Debug,
{
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "MapSheet (")?;
        writeln!(f, "    record {:?}", self.record)?;
        writeln!(f, "    load_size {:?}", self.load_size)?;
        writeln!(f, "    data {:?}", self.data)?;
        writeln!(f, "    primary {:?}", self.primary_index)?;
        writeln!(f, "    indexes {:?}", self.indexes.len())?;
        Ok(())
    }
}

#[allow(dead_code)]
impl<'a, R, K, V> MapSheet<'a, R, K, V>
where
    R: Record<K, V> + Debug,
    K: Ord + Clone + Debug,
    V: Debug,
{
    pub fn new(record: R) -> Self {
        Self {
            record,
            load_size: 0,
            primary_index: Default::default(),
            data: Default::default(),
            indexes: vec![],
        }
    }

    pub fn add_index(&mut self, index: &'a RefCell<dyn IndexBackend<V>>) {
        self.indexes.push(index);

        let index = self.indexes.last().unwrap();

        // Update all other indexes.
        for (idx, val) in self.data.iter().enumerate() {
            if let Some(val) = val {
                index.borrow_mut().insert(val, idx);
            }
        }
    }

    pub fn store(&mut self, sheet: &mut Sheet) {
        // store header
        if let Some(header) = self.record.def_header() {
            for (col, hd) in header.iter().enumerate() {
                sheet.set_value(0, col as ucell, *hd);
            }
        }

        // header data
        let row = if self.record.def_header().is_none() {
            0
        } else {
            1
        };
        // create a view into the sheet. adjust for a header.
        let mut view = SheetView::new(sheet, row, 0);

        let mut row = 0u32;
        for (_key, idx) in &self.primary_index {
            let val = self.data.get(*idx);
            // maybe not found
            if let Some(val) = val {
                // or removed
                if let Some(val) = val {
                    self.record.store(&mut view, row, val);
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
    }

    pub fn load(&mut self, sheet: &mut Sheet) {
        // reset
        self.data = Default::default();
        self.primary_index = Default::default();
        self.load_size = 0;

        // header data
        let row = if self.record.def_header().is_none() {
            0
        } else {
            1
        };

        // create a view into the sheet. ignore the header, if any.
        let mut view = SheetView::new(sheet, row, 0);

        let mut row = 0;
        loop {
            let val = self.record.load(&mut view, row);
            println!("{:?}", val);

            if let Some(val) = val {
                let key = self.record.primary(&val);
                self.primary_index.insert(key, row as usize);
                self.data.push(Some(val));
            } else {
                break;
            }

            row += 1;
        }

        // Update all other indexes.
        for (idx, val) in self.data.iter().enumerate() {
            if let Some(val) = val {
                for index in &self.indexes {
                    index.borrow_mut().insert(val, idx);
                }
            }
        }

        self.load_size = self.data.len();
    }
}

#[cfg(test)]
mod tests {
    use crate::mapstruct::{link_index, Index2, IndexValueToKey, MapSheet, Record, SheetView};
    use crate::{Sheet, WorkBook};

    #[derive(Debug, Default)]
    pub struct Artikel {
        pub artnr: u32,
        pub artbez: String,
        pub grp1: String,
        pub vkp1: f64,
        pub ust: f64,
        pub vkeh: String,
        pub aktiv: bool,
        pub bestand: u32,
    }

    #[derive(Default, Debug)]
    pub struct ArtikelRecord {}

    impl Record<u32, Artikel> for ArtikelRecord {
        fn primary(&self, val: &Artikel) -> u32 {
            val.artnr
        }

        fn def_header(&self) -> Option<&'static [&'static str]> {
            Some(&[
                "ArtNr",
                "Bezeichnung",
                "Gruppe",
                "VK",
                "USt",
                "EH",
                "Aktiv",
                "Bestand",
            ])
        }

        fn load(&self, sheet: &SheetView, n: u32) -> Option<Artikel> {
            if let Some(artnr) = sheet.value(n, 0).as_u32_opt() {
                let mut artikel = Artikel::default();
                artikel.artnr = artnr;
                artikel.artbez = sheet.value(n, 1).as_str_or("").to_string();
                artikel.grp1 = sheet.value(n, 2).as_str_or("").to_string();
                artikel.vkp1 = sheet.value(n, 3).as_f64_or(0f64);
                artikel.ust = sheet.value(n, 4).as_f64_or(0f64);
                artikel.vkeh = sheet.value(n, 5).as_str_or("").to_string();
                artikel.aktiv = sheet.value(n, 6).as_bool_or(true);
                artikel.bestand = sheet.value(n, 7).as_u32_or(0);
                Some(artikel)
            } else {
                None
            }
        }

        fn store(&self, sheet: &mut SheetView, n: u32, val: &Artikel) {
            sheet.set_value(n, 0, val.artnr);
            sheet.set_value(n, 1, val.artbez.clone());
            sheet.set_value(n, 2, val.grp1.clone());
            sheet.set_value(n, 3, val.vkp1);
            sheet.set_value(n, 4, val.ust);
            sheet.set_value(n, 5, val.vkeh.clone());
            sheet.set_value(n, 6, val.aktiv);
            sheet.set_value(n, 7, val.bestand);
        }
    }

    #[derive(Debug)]
    pub struct ArtbezIndex {}

    impl IndexValueToKey<Artikel, String> for ArtbezIndex {
        fn key<'a>(&self, val: &'a Artikel) -> &'a String {
            &val.artbez
        }
    }

    #[derive(Debug)]
    pub struct GrpIndex {}

    impl IndexValueToKey<Artikel, String> for GrpIndex {
        fn key<'a>(&self, val: &'a Artikel) -> &'a String {
            &val.grp1
        }
    }

    #[test]
    fn test_struct() {
        let mut _book = WorkBook::new();

        let mut sheet = Sheet::new_with_name("artikel");
        for idx in 1..9 {
            sheet.set_value(idx, 0, 201 + idx);
            sheet.set_value(idx, 1, format!("sample{}", idx));
        }

        let mut map0 = MapSheet::new(ArtikelRecord::default());
        map0.load(&mut sheet);
        dbg!(&map0);

        let mut idx0 = Index2::new(ArtbezIndex {});
        map0.add_index(&idx0);

        let mut idx1 = Index2::new(GrpIndex {});
        link_index(&idx1, &mut map0);

        idx0.borrow().find("sample".to_string());

        dbg!(&map0);
        dbg!(&idx0);
        dbg!(&idx1);
    }
}
