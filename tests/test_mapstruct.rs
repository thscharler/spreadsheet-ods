use spreadsheet_ods::mapstruct::{ExtractKey, Index2, MapError, MapSheet, Recorder, SheetView};
use spreadsheet_ods::Sheet;

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

impl Recorder<u32, Artikel> for ArtikelRecord {
    fn primary_key(&self, val: &Artikel) -> u32 {
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

    fn load(&self, sheet: &SheetView, n: u32) -> Result<Option<Artikel>, MapError<u32, Artikel>> {
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
            Ok(Some(artikel))
        } else {
            Ok(None)
        }
    }

    fn store(
        &self,
        sheet: &mut SheetView,
        n: u32,
        val: &Artikel,
    ) -> Result<(), MapError<u32, Artikel>> {
        sheet.set_value(n, 0, val.artnr);
        sheet.set_value(n, 1, val.artbez.clone());
        sheet.set_value(n, 2, val.grp1.clone());
        sheet.set_value(n, 3, val.vkp1);
        sheet.set_value(n, 4, val.ust);
        sheet.set_value(n, 5, val.vkeh.clone());
        sheet.set_value(n, 6, val.aktiv);
        sheet.set_value(n, 7, val.bestand);
        Ok(())
    }
}

#[derive(Debug)]
pub struct ArtbezIndex {}

impl ExtractKey<Artikel, String> for ArtbezIndex {
    fn key<'a>(&self, val: &'a Artikel) -> &'a String {
        &val.artbez
    }
}

#[derive(Debug)]
pub struct GrpIndex {}

impl ExtractKey<Artikel, String> for GrpIndex {
    fn key<'a>(&self, val: &'a Artikel) -> &'a String {
        &val.grp1
    }
}

fn load<'a>() -> Result<MapSheet<'a, u32, Artikel>, MapError<u32, Artikel>> {
    let mut sheet = Sheet::new_with_name("artikel");
    for idx in 1..9 {
        sheet.set_value(idx, 0, 201 + idx);
        sheet.set_value(idx, 1, format!("sample{}", idx));
    }

    let mut map0 = MapSheet::new(ArtikelRecord::default());
    map0.load(&mut sheet)?;

    Ok(map0)
}

#[test]
fn test_struct() -> Result<(), MapError<u32, Artikel>> {
    let idx0 = Index2::new(ArtbezIndex {});

    let mut map0 = load()?;

    map0.add_index(&idx0);
    let idx1 = Index2::new(GrpIndex {});

    map0.add_index(&idx1);

    assert_eq!(map0.len(), 8);
    assert_eq!(map0.len_vec(), 8);

    assert_eq!(map0.find_idx(&202), Some(0));
    assert_eq!(map0.find_idx(&102), None);

    assert_eq!(map0.get_idx(0).unwrap().artnr, 202);
    assert!(map0.get_idx(10).is_none());
    assert_eq!(map0.get_idx_mut(0).unwrap().artnr, 202);
    assert!(map0.get_idx_mut(10).is_none());

    let v = Artikel {
        artnr: 999,
        artbez: "inserted".to_string(),
        grp1: "gr9".to_string(),
        vkp1: 0.0,
        ust: 0.0,
        vkeh: "pc".to_string(),
        aktiv: true,
        bestand: 0,
    };
    map0.insert(v)?;
    assert_eq!(map0.len(), 9);
    assert_eq!(map0.len_vec(), 9);
    let v = Artikel {
        artnr: 998,
        artbez: "inserted".to_string(),
        grp1: "gr9".to_string(),
        vkp1: 0.0,
        ust: 0.0,
        vkeh: "pc".to_string(),
        aktiv: true,
        bestand: 0,
    };
    map0.insert(v)?;
    assert_eq!(map0.len(), 10);
    assert_eq!(map0.len_vec(), 10);

    assert_eq!(map0.remove_idx(0).unwrap().artnr, 202);
    assert_eq!(map0.len(), 9);
    assert_eq!(map0.len_vec(), 10);
    assert!(map0.remove_idx(0).is_none());

    assert!(map0.get(&202).is_none());
    assert_eq!(map0.get(&203).unwrap().artnr, 203);
    assert!(map0.get_mut(&202).is_none());
    assert_eq!(map0.get_mut(&203).unwrap().artnr, 203);

    let v = Artikel {
        artnr: 1111,
        artbez: "updated".to_string(),
        grp1: "updated-gr9".to_string(),
        vkp1: 0.0,
        ust: 0.0,
        vkeh: "pc".to_string(),
        aktiv: true,
        bestand: 0,
    };
    map0.update(&998, v)?;

    dbg!(&map0);
    dbg!(&idx0);
    dbg!(&idx1);

    Ok(())
}

#[test]
fn test_double_insert() -> Result<(), MapError<u32, Artikel>> {
    let mut map0 = load()?;

    let v = Artikel {
        artnr: 998,
        artbez: "inserted".to_string(),
        grp1: "gr9".to_string(),
        vkp1: 0.0,
        ust: 0.0,
        vkeh: "pc".to_string(),
        aktiv: true,
        bestand: 0,
    };
    map0.insert(v)?;

    let v = Artikel {
        artnr: 998,
        artbez: "inserted".to_string(),
        grp1: "gr9".to_string(),
        vkp1: 0.0,
        ust: 0.0,
        vkeh: "pc".to_string(),
        aktiv: true,
        bestand: 0,
    };
    assert!(map0.insert(v).is_err());
    //map0.insert(v)?;

    Ok(())
}
