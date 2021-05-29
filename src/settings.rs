use chrono::NaiveDateTime;
use std::collections::HashMap;

#[derive(Debug, Clone, PartialOrd, PartialEq)]
pub enum ConfigValue {
    Base64Binary(String),
    Boolean(bool),
    DateTime(chrono::NaiveDateTime),
    Double(f64),
    Int(i32),
    Long(i64),
    Short(i16),
    String(String),
}

impl ConfigValue {}

impl From<&str> for ConfigValue {
    fn from(v: &str) -> Self {
        ConfigValue::String(v.into())
    }
}

impl From<String> for ConfigValue {
    fn from(v: String) -> Self {
        ConfigValue::String(v)
    }
}

impl From<bool> for ConfigValue {
    fn from(v: bool) -> Self {
        ConfigValue::Boolean(v)
    }
}

impl From<NaiveDateTime> for ConfigValue {
    fn from(v: NaiveDateTime) -> Self {
        ConfigValue::DateTime(v)
    }
}

impl From<f64> for ConfigValue {
    fn from(v: f64) -> Self {
        ConfigValue::Double(v)
    }
}

impl From<i16> for ConfigValue {
    fn from(v: i16) -> Self {
        ConfigValue::Short(v)
    }
}

impl From<i32> for ConfigValue {
    fn from(v: i32) -> Self {
        ConfigValue::Int(v)
    }
}

impl From<i64> for ConfigValue {
    fn from(v: i64) -> Self {
        ConfigValue::Long(v)
    }
}

#[derive(Debug, Clone)]
pub struct ConfigSet {
    set: HashMap<String, ConfigItem>,
}

impl Default for ConfigSet {
    fn default() -> Self {
        ConfigSet::new()
    }
}

impl ConfigSet {
    pub fn new() -> Self {
        Self {
            set: Default::default(),
        }
    }

    pub fn insert<S, V>(&mut self, name: S, item: V)
    where
        S: Into<String>,
        V: Into<ConfigItem>,
    {
        self.set.insert(name.into(), item.into());
    }

    pub fn get<S>(&self, name: S) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        self.set.get(name.as_ref())
    }
}

#[derive(Debug, Clone)]
pub struct ConfigVec {
    vec: HashMap<String, HashMap<String, ConfigItem>>,
}

impl ConfigVec {
    pub fn new() -> Self {
        Self {
            vec: HashMap::default(),
        }
    }

    pub fn insert<S, T, V>(&mut self, index: S, item_name: T, item: V)
    where
        S: Into<String>,
        T: Into<String>,
        V: Into<ConfigItem>,
    {
        self.vec
            .entry(index.into())
            .or_insert_with(HashMap::new)
            .insert(item_name.into(), item.into());
    }

    pub fn get<S, T>(&self, index: S, item_name: T) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
        T: AsRef<str>,
    {
        if let Some(map) = self.vec.get(index.as_ref()) {
            map.get(item_name.as_ref())
        } else {
            None
        }
    }

    // pub fn remove<S, T>(&mut self, index: S, item_name: T) -> Option<ConfigItem>
    // where
    //     S: AsRef<str>,
    //     T: AsRef<str>,
    // {
    //     if let Some(map) = self.vec.get_mut(index.as_ref()) {
    //         map.remove(item_name.as_ref())
    //     } else {
    //         None
    //     }
    // }
}

#[derive(Debug, Clone)]
pub struct ConfigMap {
    map: HashMap<String, HashMap<String, ConfigItem>>,
}

impl ConfigMap {
    pub fn new() -> Self {
        Self {
            map: Default::default(),
        }
    }

    pub fn insert<S, T, V>(&mut self, map_name: S, item_name: T, item: V)
    where
        S: Into<String>,
        T: Into<String>,
        V: Into<ConfigItem>,
    {
        self.map
            .entry(map_name.into())
            .or_insert_with(HashMap::new)
            .insert(item_name.into(), item.into());
    }

    pub fn get<S, T>(&self, map_name: S, item_name: T) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
        T: AsRef<str>,
    {
        if let Some(map) = self.map.get(map_name.as_ref()) {
            map.get(item_name.as_ref())
        } else {
            None
        }
    }

    // pub fn remove<S, T>(&mut self, map_name: S, item_name: T) -> Option<ConfigItem>
    // where
    //     S: AsRef<str>,
    //     T: AsRef<str>,
    // {
    //     if let Some(map) = self.map.get_mut(map_name.as_ref()) {
    //         map.remove(item_name.as_ref())
    //     } else {
    //         None
    //     }
    // }
}

#[derive(Debug, Clone)]
pub struct ConfigEntry {
    set: HashMap<String, ConfigItem>,
}

impl Default for ConfigEntry {
    fn default() -> Self {
        ConfigEntry::new()
    }
}

impl ConfigEntry {
    pub fn new() -> Self {
        Self {
            set: Default::default(),
        }
    }

    pub fn insert<S, V>(&mut self, name: S, item: V)
    where
        S: Into<String>,
        V: Into<ConfigItem>,
    {
        self.set.insert(name.into(), item.into());
    }

    pub fn get<S>(&self, name: S) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        self.set.get(name.as_ref())
    }
}

#[derive(Debug, Clone)]
pub enum ConfigItem {
    Value(ConfigValue),
    Set(ConfigSet),
    Vec(ConfigVec),
    Map(ConfigMap),
}

impl<T> From<T> for ConfigItem
where
    ConfigValue: From<T>,
{
    fn from(v: T) -> Self {
        ConfigItem::Value(ConfigValue::from(v))
    }
}

impl From<ConfigSet> for ConfigItem {
    fn from(v: ConfigSet) -> Self {
        ConfigItem::Set(v)
    }
}

impl From<ConfigVec> for ConfigItem {
    fn from(v: ConfigVec) -> Self {
        ConfigItem::Vec(v)
    }
}

impl From<ConfigMap> for ConfigItem {
    fn from(v: ConfigMap) -> Self {
        ConfigItem::Map(v)
    }
}

impl ConfigItem {
    pub fn get<S>(&self, names: &[S]) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        match self {
            ConfigItem::Value(_) => {
                // no deeper nesting, ok
                if names.is_empty() {
                    Some(self)
                } else {
                    None
                }
            }
            ConfigItem::Set(set) => {
                if let Some((name, rest)) = names.split_first() {
                    if let Some(v) = set.get(name) {
                        v.get(rest)
                    } else {
                        None
                    }
                } else {
                    Some(self)
                }
            }
            ConfigItem::Vec(vec) => {
                if let Some((index, rest)) = names.split_first() {
                    if let Some((item_name, rest)) = rest.split_first() {
                        if let Some(v) = vec.get(index, item_name) {
                            v.get(rest)
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                } else {
                    Some(self)
                }
            }
            ConfigItem::Map(map) => {
                if let Some((map_name, rest)) = names.split_first() {
                    dbg!(map_name.as_ref());
                    if let Some((item_name, rest)) = rest.split_first() {
                        dbg!(item_name.as_ref());
                        if let Some(v) = map.get(map_name, item_name) {
                            dbg!(v);
                            v.get(rest)
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                } else {
                    Some(self)
                }
            }
        }
    }

    pub fn get_value<S>(&self, names: &[S]) -> Option<&ConfigValue>
    where
        S: AsRef<str>,
    {
        match self {
            ConfigItem::Value(ref v) => {
                // no deeper nesting, ok
                if names.is_empty() {
                    Some(v)
                } else {
                    None
                }
            }
            ConfigItem::Set(set) => {
                if let Some((name, rest)) = names.split_first() {
                    if let Some(v) = set.get(name) {
                        v.get_value(rest)
                    } else {
                        None
                    }
                } else {
                    None
                }
            }
            ConfigItem::Vec(vec) => {
                if let Some((index, rest)) = names.split_first() {
                    if let Some((item_name, rest)) = rest.split_first() {
                        if let Some(v) = vec.get(index, item_name) {
                            v.get_value(rest)
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                } else {
                    None
                }
            }
            ConfigItem::Map(map) => {
                if let Some((map_name, rest)) = names.split_first() {
                    dbg!(map_name.as_ref());
                    if let Some((item_name, rest)) = rest.split_first() {
                        dbg!(item_name.as_ref());
                        if let Some(v) = map.get(map_name, item_name) {
                            dbg!(v);
                            v.get_value(rest)
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                } else {
                    None
                }
            }
        }
    }
}

#[derive(Debug, Clone)]
pub struct Config {
    config: ConfigSet,
}

impl Default for Config {
    fn default() -> Self {
        Config::new()
    }
}

impl Config {
    pub fn new() -> Self {
        Self {
            config: Default::default(),
        }
    }

    pub fn insert<S, V>(&mut self, name: S, item: V)
    where
        S: Into<String>,
        V: Into<ConfigItem>,
    {
        self.config.insert(name.into(), item.into());
    }

    // pub fn get<S>(&self, name: &S) -> Option<&ConfigItem>
    // where
    //     S: AsRef<str>,
    // {
    //     self.config.get(name)
    // }

    // pub fn remove<S>(&mut self, name: &S) -> Option<ConfigItem>
    // where
    //     S: AsRef<str>,
    // {
    //     self.config.remove(name)
    // }

    pub fn get<S>(&self, names: &[S]) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        if let Some((name, rest)) = names.split_first() {
            if let Some(set) = self.config.get(name) {
                set.get(rest)
            } else {
                None
            }
        } else {
            None
        }
    }

    pub fn get_value<S>(&self, names: &[S]) -> Option<&ConfigValue>
    where
        S: AsRef<str>,
    {
        if let Some((name, rest)) = names.split_first() {
            if let Some(set) = self.config.get(name) {
                set.get_value(rest)
            } else {
                None
            }
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::settings::{Config, ConfigMap, ConfigSet, ConfigValue};

    #[test]
    fn test_config() {
        let mut config = Config::new();

        let mut view_settings = ConfigSet::new();
        view_settings.insert("VisibleAreaTop", 903);
        config.insert("ooo:view-settings", view_settings);

        let mut configuration_settings = ConfigSet::new();
        configuration_settings.insert("HasSheetTabs".to_string(), true);
        configuration_settings.insert("ShowNotes", true);
        configuration_settings.insert("GridColor", 12632256);
        configuration_settings.insert("LinkUpdateMode", 3i16);
        configuration_settings.insert("PrinterSetup", ConfigValue::base64("unknowgarbage"));

        let mut script_configuration = ConfigMap::new();
        script_configuration.insert("Tabelle1", "CodeName", "Tabelle1");
        configuration_settings.insert("ScriptConfiguration", script_configuration);

        config.insert("ooo:configuration-settings", configuration_settings);

        assert_eq!(config.get_value(&["ooo:view-settings", "ShowNotes"]), None);
        assert_eq!(
            config.get_value_or(
                &["ooo:view-settings", "ShowNotes"],
                &ConfigValue::from(true)
            ),
            &ConfigValue::Boolean(true)
        );
        assert_eq!(
            config.get_value(&["ooo:view-settings", "VisibleAreaTop"]),
            Some(&ConfigValue::Int(903))
        );
        assert_eq!(
            config.get_value(&["ooo:configuration-settings", "ShowNotes"]),
            Some(&ConfigValue::Boolean(true))
        );
        assert_eq!(
            config.get_value(&[
                "ooo:configuration-settings",
                "ScriptConfiguration",
                "Tabelle1",
                "CodeName"
            ]),
            Some(&ConfigValue::String("Tabelle1".to_string()))
        );
    }
}
