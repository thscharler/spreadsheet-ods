macro_rules! styles_styles {
    ($style:ident, $styleref:ident) => {
        impl $style {
            /// Origin of the style, either styles.xml oder content.xml
            pub fn origin(&self) -> StyleOrigin {
                self.origin
            }

            /// Changes the origin.
            pub fn set_origin(&mut self, origin: StyleOrigin) {
                self.origin = origin;
            }

            /// Usage for the style.
            pub fn styleuse(&self) -> StyleUse {
                self.styleuse
            }

            /// Usage for the style.
            pub fn set_styleuse(&mut self, styleuse: StyleUse) {
                self.styleuse = styleuse;
            }

            /// Stylename
            pub fn name(&self) -> &str {
                &self.name
            }

            /// Stylename
            pub fn set_name<S: Into<String>>(&mut self, name: S) {
                self.name = name.into();
            }

            /// Returns the name as a style reference.
            pub fn style_ref(&self) -> $styleref {
                $styleref::from(self.name())
            }

            style_auto_update!(attr);
            style_class!(attr);
            style_display_name!(attr);
            style_parent_style_name!(attr, $styleref);
        }
    };
}

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

        impl AsRef<$l> for $l {
            fn as_ref(&self) -> &$l {
                self
            }
        }

        impl Display for $l {
            fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
                write!(f, "{}", self.name)
            }
        }

        impl $l {
            /// Reference as str.
            pub fn as_str(&self) -> &str {
                self.name.as_str()
            }
        }
    };
}

macro_rules! xml_id {
    ($acc:ident) => {
        /// The table:end-y attribute specifies the y-coordinate of the end position of a shape relative to
        /// the top-left edge of a cell. The size attributes of the shape are ignored.
        pub fn set_xml_id<S: Into<String>>(&mut self, id: S) {
            self.$acc.set_attr("xml_id", id.into());
        }
    };
}
