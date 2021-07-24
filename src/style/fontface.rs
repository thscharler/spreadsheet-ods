use crate::attrmap2::AttrMap2;
use crate::style::units::FontPitch;
use crate::style::StyleOrigin;

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontFaceDecl {
    name: String,
    /// From where did we get this style.
    origin: StyleOrigin,
    /// All other attributes.
    attr: AttrMap2,
}

impl FontFaceDecl {
    /// New, empty.
    pub fn new() -> Self {
        Self {
            name: "".to_string(),
            origin: Default::default(),
            attr: Default::default(),
        }
    }

    /// New, with a name.
    pub fn new_with_name<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            origin: StyleOrigin::Content,
            attr: Default::default(),
        }
    }

    /// Set the name.
    pub fn set_name<V: Into<String>>(&mut self, name: V) {
        self.name = name.into();
    }

    /// Returns the name.
    pub fn name(&self) -> &String {
        &self.name
    }

    /// Origin of the style
    pub fn set_origin(&mut self, origin: StyleOrigin) {
        self.origin = origin;
    }

    /// Origin of the style
    pub fn origin(&self) -> StyleOrigin {
        self.origin
    }

    /// General attributes.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    font_decl!(attrmap_mut);
}
