use std::collections::HashMap;

use chrono::NaiveDateTime;

use crate::ucell;

/// The possible value types for the configuration.
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

impl From<ucell> for ConfigValue {
    fn from(v: ucell) -> Self {
        ConfigValue::Int(v as i32)
    }
}

impl From<i64> for ConfigValue {
    fn from(v: i64) -> Self {
        ConfigValue::Long(v)
    }
}

#[derive(Debug, Clone, Copy)]
pub enum ConfigItemType {
    Value,
    Set,
    Vec,
    Map,
    Entry,
}

impl From<&ConfigItem> for ConfigItemType {
    fn from(item: &ConfigItem) -> Self {
        match item {
            ConfigItem::Value(_) => ConfigItemType::Value,
            ConfigItem::Set(_) => ConfigItemType::Set,
            ConfigItem::Vec(_) => ConfigItemType::Vec,
            ConfigItem::Map(_) => ConfigItemType::Map,
            ConfigItem::Entry(_) => ConfigItemType::Entry,
        }
    }
}

impl From<&mut ConfigItem> for ConfigItemType {
    fn from(item: &mut ConfigItem) -> Self {
        match item {
            ConfigItem::Value(_) => ConfigItemType::Value,
            ConfigItem::Set(_) => ConfigItemType::Set,
            ConfigItem::Vec(_) => ConfigItemType::Vec,
            ConfigItem::Map(_) => ConfigItemType::Map,
            ConfigItem::Entry(_) => ConfigItemType::Entry,
        }
    }
}

impl PartialEq<ConfigItem> for ConfigItemType {
    fn eq(&self, other: &ConfigItem) -> bool {
        other == self
    }
}

impl PartialEq<ConfigItemType> for ConfigItem {
    fn eq(&self, other: &ConfigItemType) -> bool {
        match self {
            ConfigItem::Value(_) => match other {
                ConfigItemType::Value => true,
                _ => false,
            },
            ConfigItem::Set(_) => match other {
                ConfigItemType::Set => true,
                _ => false,
            },
            ConfigItem::Vec(_) => match other {
                ConfigItemType::Vec => true,
                _ => false,
            },
            ConfigItem::Map(_) => match other {
                ConfigItemType::Map => true,
                _ => false,
            },
            ConfigItem::Entry(_) => match other {
                ConfigItemType::Entry => true,
                _ => false,
            },
        }
    }
}

/// Unifies values and sets of values. The branch structure of the tree.
#[derive(Debug, Clone, PartialEq)]
pub enum ConfigItem {
    Value(ConfigValue),
    Set(HashMap<String, ConfigItem>),
    Vec(HashMap<String, ConfigItem>),
    Map(HashMap<String, ConfigItem>),
    Entry(HashMap<String, ConfigItem>),
}

/// Nice conversion for everything that can be converted to a ConfigValue
/// can directly be converted to a ConfigItem too.
impl<T> From<T> for ConfigItem
where
    ConfigValue: From<T>,
{
    fn from(v: T) -> Self {
        ConfigItem::Value(ConfigValue::from(v))
    }
}

impl Default for ConfigItem {
    fn default() -> Self {
        ConfigItem::new_set()
    }
}

impl ConfigItem {
    pub fn new(itype: ConfigItemType) -> Self {
        match itype {
            ConfigItemType::Value => panic!("new with type works only for map-types"),
            ConfigItemType::Set => ConfigItem::Set(HashMap::new()),
            ConfigItemType::Vec => ConfigItem::Vec(HashMap::new()),
            ConfigItemType::Map => ConfigItem::Map(HashMap::new()),
            ConfigItemType::Entry => ConfigItem::Entry(HashMap::new()),
        }
    }

    pub fn new_set() -> Self {
        Self::Set(HashMap::new())
    }

    pub fn new_vec() -> Self {
        Self::Vec(HashMap::new())
    }

    pub fn new_map() -> Self {
        Self::Map(HashMap::new())
    }

    pub fn new_entry() -> Self {
        Self::Entry(HashMap::new())
    }

    fn as_value(&self) -> Option<&ConfigValue> {
        match self {
            ConfigItem::Value(v) => Some(v),
            ConfigItem::Set(_) => None,
            ConfigItem::Vec(_) => None,
            ConfigItem::Map(_) => None,
            ConfigItem::Entry(_) => None,
        }
    }

    fn is_map(&self) -> bool {
        match self {
            ConfigItem::Value(_) => false,
            ConfigItem::Set(_) => true,
            ConfigItem::Vec(_) => true,
            ConfigItem::Map(_) => true,
            ConfigItem::Entry(_) => true,
        }
    }

    fn as_map(&self) -> Option<&HashMap<String, ConfigItem>> {
        match self {
            ConfigItem::Value(_) => None,
            ConfigItem::Set(m) => Some(m),
            ConfigItem::Vec(m) => Some(m),
            ConfigItem::Map(m) => Some(m),
            ConfigItem::Entry(m) => Some(m),
        }
    }

    fn as_map_mut(&mut self) -> Option<&mut HashMap<String, ConfigItem>> {
        match self {
            ConfigItem::Value(_) => None,
            ConfigItem::Set(m) => Some(m),
            ConfigItem::Vec(m) => Some(m),
            ConfigItem::Map(m) => Some(m),
            ConfigItem::Entry(m) => Some(m),
        }
    }

    /// Iterate over (k,v) pairs.
    pub fn iter(&self) -> ConfigIter {
        if let Some(m) = self.as_map() {
            ConfigIter { it: Some(m.iter()) }
        } else {
            ConfigIter { it: None }
        }
    }

    /// Adds a new ConfigItem
    pub fn insert<S, V>(&mut self, name: S, item: V)
    where
        S: Into<String>,
        V: Into<ConfigItem>,
    {
        if let Some(m) = self.as_map_mut() {
            m.insert(name.into(), item.into());
        } else {
            panic!();
        }
    }

    /// Returns a ConfigItem
    pub fn get<S>(&self, name: S) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        if let Some(m) = self.as_map() {
            m.get(name.as_ref())
        } else {
            panic!()
        }
    }

    /// Recursive get for any ConfigItem.
    pub fn create_path<S>(&mut self, names: &[(S, ConfigItemType)]) -> &mut ConfigItem
    where
        S: AsRef<str>,
    {
        if self.is_map() {
            // some name parts left?
            if let Some(((name, itype), rest)) = names.split_first() {
                // create if non existent
                let item = self
                    .as_map_mut()
                    .unwrap()
                    .entry(name.as_ref().to_string())
                    .or_insert_with(|| ConfigItem::new(*itype));

                if !(item == itype) {
                    // close, but not good enough
                    panic!(
                        "types don't match {:?} <> {:?}",
                        ConfigItemType::from(item),
                        itype
                    );
                } else {
                    // recurse
                    item.create_path(rest)
                }
            } else {
                // last path element is what we want
                self
            }
        } else {
            // not a map
            panic!("path ends in a value");
        }
    }

    /// Recursive get for any ConfigItem.
    pub fn get_rec<S>(&self, names: &[S]) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        if let Some(map) = self.as_map() {
            if let Some((name, rest)) = names.split_first() {
                if let Some(item) = map.get(name.as_ref()) {
                    item.get_rec(rest)
                } else {
                    None
                }
            } else {
                Some(self)
            }
        } else {
            // no deeper nesting, ok
            if names.is_empty() {
                Some(self)
            } else {
                None
            }
        }
    }

    /// Recursive get for only the ConfigValue leaves.
    pub fn get_value_rec<S>(&self, names: &[S]) -> Option<&ConfigValue>
    where
        S: AsRef<str>,
    {
        if let Some(map) = self.as_map() {
            if let Some((name, rest)) = names.split_first() {
                if let Some(item) = map.get(name.as_ref()) {
                    item.get_value_rec(rest)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            // no deeper nesting, ok
            if names.is_empty() {
                self.as_value()
            } else {
                None
            }
        }
    }
}

pub struct ConfigIter<'a> {
    it: Option<std::collections::hash_map::Iter<'a, String, ConfigItem>>,
}

impl<'a> Iterator for ConfigIter<'a> {
    type Item = (&'a String, &'a ConfigItem);

    fn next(&mut self) -> Option<Self::Item> {
        if let Some(it) = &mut self.it {
            it.next()
        } else {
            None
        }
    }
}

/// Basic wrapper around a ConfigSet. Root of the config tree.
#[derive(Debug, Clone)]
pub struct Config {
    config: ConfigItem,
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

    /// Iterate over the (k,v) pairs.
    pub fn iter(&self) -> ConfigIter {
        self.config.iter()
    }

    /// Add an item.
    pub fn insert<S, V>(&mut self, name: S, item: V)
    where
        S: Into<String>,
        V: Into<ConfigItem>,
    {
        self.config.insert(name.into(), item.into());
    }

    /// Recursive get.
    pub fn get<S>(&self, names: &[S]) -> Option<&ConfigItem>
    where
        S: AsRef<str>,
    {
        self.config.get_rec(names)
    }

    /// Recursive get, only for ConfigValue leaves.
    pub fn get_value<S>(&self, names: &[S]) -> Option<&ConfigValue>
    where
        S: AsRef<str>,
    {
        self.config.get_value_rec(names)
    }

    pub fn create_path<S>(&mut self, names: &[(S, ConfigItemType)]) -> &mut ConfigItem
    where
        S: AsRef<str>,
    {
        self.config.create_path(names)
    }
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use crate::settings::{Config, ConfigItem, ConfigItemType, ConfigValue};

    fn setup_config() -> Config {
        let mut config = Config::new();
        {
            let mut view_settings = ConfigItem::new_set();
            view_settings.insert("VisibleAreaTop", 903);
            config.insert("ooo:view-settings", view_settings);
        }
        {
            let mut configuration_settings = ConfigItem::new_set();
            configuration_settings.insert("HasSheetTabs".to_string(), true);
            configuration_settings.insert("ShowNotes", true);
            configuration_settings.insert("GridColor", 12632256);
            configuration_settings.insert("LinkUpdateMode", 3i16);
            configuration_settings.insert(
                "PrinterSetup",
                ConfigValue::Base64Binary("unknown_garbage".to_string()),
            );
            {
                let mut script_configuration = ConfigItem::new_map();
                {
                    let mut tabelle1 = ConfigItem::new_entry();
                    tabelle1.insert("CodeName", "Tabelle1");
                    script_configuration.insert("Tabelle1", tabelle1);
                }
                configuration_settings.insert("ScriptConfiguration", script_configuration);
            }
            config.insert("ooo:configuration-settings", configuration_settings);
        }

        config
    }

    #[test]
    fn test_config() {
        let mut config = setup_config();

        assert_eq!(config.get_value(&["ooo:view-settings", "ShowNotes"]), None);
        assert_eq!(config.get_value(&["ooo:view-settings", "ShowNotes"]), None);
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

        let v = config.create_path(&[
            ("ooo:configuration-settings", ConfigItemType::Set),
            ("ScriptConfiguration", ConfigItemType::Map),
            ("Tabelle2", ConfigItemType::Entry),
        ]);
        assert_eq!(v, &ConfigItem::Entry(HashMap::new()));
    }

    #[test]
    #[should_panic]
    fn test_create_path() {
        let mut config = setup_config();
        let _v = config.create_path(&[
            ("ooo:configuration-settings", ConfigItemType::Map), // here
            ("ScriptConfiguration", ConfigItemType::Map),
            ("Tabelle2", ConfigItemType::Entry),
        ]);
    }

    #[test]
    #[should_panic]
    fn test_create_path2() {
        let mut config = setup_config();
        let _v = config.create_path(&[
            ("ooo:configuration-settings", ConfigItemType::Set),
            ("ScriptConfiguration", ConfigItemType::Map),
            ("Tabelle1", ConfigItemType::Entry),
            ("CodeName", ConfigItemType::Value), // here
        ]);
    }
}
