use spreadsheet_ods::mapstruct::{MapError, MapSheet, Recorder, SheetView};
use spreadsheet_ods::Value;

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

pub struct ArtikelRecorder {}

impl Recorder<u32, Artikel> for ArtikelRecorder {
    fn primary_key(&self, val: &Artikel) -> u32 {
        val.artnr
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

pub struct ArtikelDB<'a> {
    pub artikel: MapSheet<'a, u32, Artikel>,
}

#[test]
fn test_artikel() {}
