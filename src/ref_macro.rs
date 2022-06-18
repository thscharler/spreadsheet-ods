/// Generates a name reference for a style.
macro_rules! style_ref {
    ($l:ident) => {
        /// Reference
        #[derive(Debug, Clone)]
        pub struct $l {
            name: String,
        }

        impl From<String> for $l {
            fn from(name: String) -> Self {
                Self { name }
            }
        }

        impl From<&String> for $l {
            fn from(name: &String) -> Self {
                Self {
                    name: name.to_string(),
                }
            }
        }

        impl From<&str> for $l {
            fn from(name: &str) -> Self {
                Self {
                    name: name.to_string(),
                }
            }
        }

        impl From<$l> for String {
            fn from(name: $l) -> Self {
                name.to_string()
            }
        }

        impl $l {
            /// Reference as str.
            pub fn as_str(&self) -> &str {
                self.name.as_str()
            }

            /// Reference as String.
            pub fn to_string(&self) -> String {
                self.name.clone()
            }
        }
    };
}

/// Generates a name reference for a style.
macro_rules! text_tag {
    ($tag:ident, $xml:literal) => {
        /// $literal
        #[derive(Debug)]
        pub struct $tag {
            xml: XmlTag,
        }

        impl From<$tag> for XmlTag {
            fn from(t: $tag) -> XmlTag {
                t.xml
            }
        }

        impl Display for $tag {
            fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), std::fmt::Error> {
                write!(f, "{}", self.xml)
            }
        }

        impl Default for $tag {
            fn default() -> Self {
                Self::new()
            }
        }

        impl $tag {
            /// Creates a new $xml.
            pub fn new() -> Self {
                $tag {
                    xml: XmlTag::new($xml),
                }
            }

            /// Appends a tag.
            pub fn tag<T: Into<XmlTag>>(mut self, tag: T) -> Self {
                self.xml.add_tag(tag);
                self
            }

            /// Appends text.
            pub fn text<S: Into<String>>(mut self, text: S) -> Self {
                self.xml.add_text(text);
                self
            }

            /// Extracts the finished XmlTag.
            pub fn into_xmltag(self) -> XmlTag {
                self.xml
            }
        }
    };
}
