use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;

use spreadsheet_ods::defaultstyles::{create_default_styles, DefaultFormat};
use spreadsheet_ods::mapstruct::{ExtractKey, Index2, MapError, MapSheet, Recorder, SheetView};
use spreadsheet_ods::{read_ods, write_ods, CellStyle, OdsError, Sheet, Value, WorkBook};

#[derive(Clone, Debug, Default)]
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

#[derive(Clone)]
pub struct ArtikelRecorder {}

impl Recorder<u32, Artikel> for ArtikelRecorder {
    fn key<'a>(&self, val: &'a Artikel) -> &'a u32 {
        &val.artnr
    }

    fn def_header(&self) -> Option<&'static [&'static str]> {
        Some(&[
            "Artikel",
            "Bezeichnung",
            "Gruppe",
            "VK",
            "USt",
            "EH",
            "Aktiv",
            "Bestand",
        ])
    }

    fn load(&self, sheet: &SheetView, row: u32) -> Result<Option<Artikel>, MapError<u32, Artikel>> {
        if matches!(sheet.value(row, 0), Value::Empty) {
            Ok(Some(Artikel {
                artnr: sheet.value(row, 0).as_u32_or(0),
                artbez: sheet.value(row, 1).as_str_or("").to_string(),
                grp1: sheet.value(row, 2).as_str_or("").to_string(),
                vkp1: sheet.value(row, 3).as_f64_or(0.0),
                ust: sheet.value(row, 4).as_f64_or(0.0),
                vkeh: sheet.value(row, 5).as_str_or("").to_string(),
                aktiv: sheet.value(row, 6).as_bool_or(true),
                bestand: sheet.value(row, 7).as_u32_or(0),
            }))
        } else {
            Ok(None)
        }
    }

    fn store(
        &self,
        sheet: &mut SheetView,
        row: u32,
        artikel: &Artikel,
    ) -> Result<(), MapError<u32, Artikel>> {
        sheet.set_value(row, 0, artikel.artnr);
        sheet.set_value(row, 1, artikel.artbez.clone());
        sheet.set_value(row, 2, artikel.grp1.clone());
        sheet.set_value(row, 3, artikel.vkp1);
        sheet.set_value(row, 4, artikel.ust);
        sheet.set_value(row, 5, artikel.vkeh.clone());
        sheet.set_value(row, 6, artikel.aktiv);
        sheet.set_value(row, 7, artikel.bestand);
        Ok(())
    }
}

pub struct IndexArtbez {}

impl ExtractKey<String, Artikel> for IndexArtbez {
    fn key<'a>(&self, val: &'a Artikel) -> &'a String {
        &val.artbez
    }
}

pub struct ArtikelDB {
    pub ods: Option<WorkBook>,
    pub artikel: MapSheet<u32, Artikel>,
    pub idx_artbez: Rc<RefCell<Index2<String, Artikel>>>,
}

impl ArtikelDB {
    pub fn new() -> Self {
        let mut adb = ArtikelDB {
            ods: None,
            artikel: MapSheet::new(ArtikelRecorder {}),
            idx_artbez: Index2::new(IndexArtbez {}),
        };

        if matches!(adb.artikel.add_index(adb.idx_artbez.clone()), Err(_)) {
            unreachable!()
        }
        adb
    }

    pub fn insert(&mut self, artikel: Artikel) -> Result<(), OdsError> {
        self.artikel.insert(artikel)?;
        Ok(())
    }

    pub fn update(&mut self, artnr_old: u32, artikel: Artikel) -> Result<(), OdsError> {
        self.artikel.update(&artnr_old, artikel)?;
        Ok(())
    }

    pub fn find_artnr(&self, artnr: u32) -> Option<&Artikel> {
        self.artikel.get(&artnr)
    }

    pub fn find_artnr_mut(&mut self, artnr: u32) -> Option<&mut Artikel> {
        self.artikel.get_mut(&artnr)
    }

    pub fn find_artbez(&self, artbez: &String) -> Option<&Artikel> {
        if let Some(idx) = self.idx_artbez.borrow().find_idx(artbez) {
            // todo: find first ok?
            if let Some(idx) = idx.iter().next() {
                return self.artikel.get_idx(*idx);
            }
        }
        None
    }

    pub fn find_artbez_mut(&mut self, artbez: &String) -> Option<&mut Artikel> {
        if let Some(idx) = self.idx_artbez.borrow().find_idx(artbez) {
            // todo: find first ok?
            if let Some(idx) = idx.iter().next() {
                return self.artikel.get_idx_mut(*idx);
            }
        }
        None
    }

    pub fn read(&mut self, file: &Path) -> Result<(), OdsError> {
        if file.exists() {
            self.ods = Some(read_ods(file)?);
            let workbook = self.ods.as_mut().unwrap();
            let sheet = workbook.sheet_mut(0);
            self.artikel.load(sheet)?;
        }
        Ok(())
    }

    pub fn write(&mut self, file: &Path) -> Result<(), OdsError> {
        if self.ods.is_none() {
            let mut workbook = WorkBook::new();
            let mut style = CellStyle::new("title1", &DefaultFormat::default());
            style.set_font_bold();
            workbook.add_cellstyle(style);
            create_default_styles(&mut workbook);
            workbook.push_sheet(Sheet::new_with_name("Artikel"));
            self.ods = Some(workbook);
        };

        let workbook = self.ods.as_mut().unwrap();
        let sheet = workbook.sheet_mut(0);

        self.artikel.store(sheet)?;

        write_ods(workbook, file)?;

        Ok(())
    }
}

#[test]
fn test_artikel() {}
