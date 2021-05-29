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

#[derive(Debug, Clone, Copy)]
pub enum ConfigSetType {
    Set,
    Vec,
    Map,
    Entry,
}

#[derive(Debug, Clone)]
pub struct ConfigSet {
    stype: ConfigSetType,
    set: HashMap<String, ConfigItem>,
}

impl Default for ConfigSet {
    fn default() -> Self {
        ConfigSet::new(ConfigSetType::Set)
    }
}

impl ConfigSet {
    pub fn new(t: ConfigSetType) -> Self {
        Self {
            stype: t,
            set: Default::default(),
        }
    }

    pub fn new_set() -> Self {
        Self {
            stype: ConfigSetType::Set,
            set: Default::default(),
        }
    }

    pub fn new_vec() -> Self {
        Self {
            stype: ConfigSetType::Vec,
            set: Default::default(),
        }
    }

    pub fn new_map() -> Self {
        Self {
            stype: ConfigSetType::Map,
            set: Default::default(),
        }
    }

    pub fn new_entry() -> Self {
        Self {
            stype: ConfigSetType::Entry,
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
