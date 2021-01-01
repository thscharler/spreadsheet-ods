use crate::attrmap::{AttrMap, AttrMapIter, AttrMapType};
use crate::sealed::Sealed;
use crate::style::attr::AttrFontDecl;
use crate::style::StyleOrigin;

/// Font declarations.
#[derive(Clone, Debug, Default)]
pub struct FontFaceDecl {
    name: String,
    /// From where did we get this style.
    origin: StyleOrigin,
    /// All other attributes.
    attr: AttrMapType,
}

impl Sealed for FontFaceDecl {}

impl AttrFontDecl for FontFaceDecl {}

impl AttrMap for FontFaceDecl {
    fn attr_map(&self) -> &AttrMapType {
        &self.attr
    }

    fn attr_map_mut(&mut self) -> &mut AttrMapType {
        &mut self.attr
    }
}

impl FontFaceDecl {
    /// New, empty.
    pub fn new() -> Self {
        Self {
            name: "".to_string(),
            origin: Default::default(),
            attr: None,
        }
    }

    /// New, with a name.
    pub fn new_with_name<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            origin: StyleOrigin::Content,
            attr: None,
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

    /// Iterator over the attributes of this pagelayout.
    pub fn attr_iter(&self) -> AttrMapIter {
        AttrMapIter::from(self.attr_map())
    }
}
