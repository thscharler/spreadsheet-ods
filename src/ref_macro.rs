/// Generates a name reference for a style.
macro_rules! style_ref {
    ($l:ident) => {
        /// Refers to a $l.
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

        impl $l {
            pub fn as_str(&self) -> &str {
                self.name.as_str()
            }

            pub fn to_string(&self) -> String {
                self.name.clone()
            }
        }
    };
}

/// Generates a name reference for a style.
macro_rules! text_tag {
    ($tag:ident, $xml:literal) => {
        pub struct $tag {
            xml: XmlTag,
        }

        impl Into<XmlTag> for $tag {
            fn into(self) -> XmlTag {
                self.xml
            }
        }

        impl $tag {
            pub fn new() -> Self {
                $tag {
                    xml: XmlTag::new($xml),
                }
            }

            pub fn tag<T: Into<XmlTag>>(mut self, tag: T) -> Self {
                self.xml.add_tag(tag);
                self
            }

            pub fn text<S: Into<String>>(mut self, text: S) -> Self {
                self.xml.add_text(text);
                self
            }
        }
    };
}
