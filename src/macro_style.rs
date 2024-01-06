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
                name.name
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

macro_rules! styles_styles2 {
    ($style:ident, $styleref:ident) => {
        impl Style for $style {
            type RefType = $styleref;

            /// Stylename
            fn name(&self) -> &str {
                &self.name
            }

            /// Stylename
            fn set_name<S: Into<String>>(&mut self, name: S) {
                self.name = name.into();
            }

            /// Returns the name as a style reference.
            fn style_ref(&self) -> $styleref {
                self.id
            }

            fn set_style_ref(&mut self, style_ref: $styleref) {
                self.id = style_ref;
            }
        }

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

            style_auto_update!(attr);
            style_class!(attr);
            style_display_name!(attr);
            style_parent_style_name!(attr, $styleref);
        }
    };
}

/// Generates a name reference for a style.
macro_rules! style_ref2 {
    ($l:ident) => {
        /// Reference
        #[derive(Debug, Clone, Copy, Hash, Eq, PartialEq, GetSize)]
        pub struct $l {
            id: u32,
        }

        impl Default for $l {
            fn default() -> Self {
                Self { id: 0 }
            }
        }

        impl From<u32> for $l {
            fn from(id: u32) -> $l {
                Self { id }
            }
        }

        impl From<Option<$l>> for $l {
            fn from(sref: Option<$l>) -> $l {
                if let Some(sref) = sref {
                    sref
                } else {
                    Default::default()
                }
            }
        }

        impl Display for $l {
            fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
                write!(f, "{}", self.id)
            }
        }

        impl StyleRef for $l {
            fn is_empty(&self) -> bool {
                self.id == 0
            }

            fn as_usize(&self) -> usize {
                self.id as usize
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
