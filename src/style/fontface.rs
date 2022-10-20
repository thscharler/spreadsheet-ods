use crate::attrmap2::AttrMap2;
use crate::style::units::{
    FontFamilyGeneric, FontPitch, FontStretch, FontStyle, FontVariant, FontWeight,
};
use crate::style::StyleOrigin;

/// The <style:font-face> element represents a font face declaration which documents the
/// properties of a font used in a document.
///
/// OpenDocument font face declarations directly correspond to the @font-face font description of
/// [CSS2] (see §15.3.1) and the <font-face> element of [SVG] (see §20.8.3).
///
/// OpenDocument font face declarations may have an unique name. This name can be used inside
/// styles (as an attribute of <style:text-properties> element) as value of the style:fontname attribute to select a font face declaration. If a font face declaration is referenced by name,
/// the font-matching algorithms for selecting a font declaration based on the font-family, font-style,
/// font-variant, font-weight and font-size descriptors are not used but the referenced font face
/// declaration is used directly. (See §15.5 [CSS2])
///
/// Consumers should implement the CSS2 font-matching algorithm with the OpenDocument font
/// face extensions. They may implement variations of the CSS2 font-matching algorithm. They may
/// implement a font-matching based only on the font face declarations, that is, a font-matching that is
/// not applied to every character independently but only once for each font face declaration. (See
/// §15.5 [CSS2])
///
/// Font face declarations support the font descriptor attributes and elements described in §20.8.3 of
/// [SVG].
#[derive(Clone, Debug, Default)]
pub struct FontFaceDecl {
    name: String,
    /// From where did we get this style.
    origin: StyleOrigin,
    /// All other attributes.
    // obsolete style:font-adornments 19.482,
    // obsolet style:font-charset 19.483,
    // ok style:font-family-generic 19.484,
    // ok style:font-pitch 19.485,
    // ok style:name 19.502,
    // ignore svg:accent-height 19.523,
    // ignore svg:alphabetic 19.524,
    // ignore svg:ascent 19.525,
    // ignore svg:bbox 19.526,
    // ignore svg:cap-height 19.527,
    // ignore svg:descent 19.531,
    // ok svg:font-family 19.532,
    // ignore svg:font-size 19.533,
    // ok svg:font-stretch 19.534,
    // ok svg:font-style 19.535,
    // ok svg:font-variant 19.536,
    // ok svg:font-weight 19.537,
    // ignore svg:hanging 19.542,
    // ignore svg:ideographic 19.544,
    // ignore svg:mathematical 19.545,
    // ignore svg:overline-position 19.549,
    // ignore svg:overline-thickness 19.550,
    // ignore svg:panose-1 19.551,
    // ignore svg:slope 19.556,
    // ignore svg:stemh 19.558,
    // ignore svg:stemv 19.559,
    // ignore svg:strikethrough-position 19.562,
    // ignore svg:strikethrough-thickness 19.563,
    // ignore svg:underline-position 19.566,
    // ignore svg:underline-thickness 19.567,
    // ignore svg:unicode-range 19.568,
    // ignore svg:unitsper-em 19.569,
    // ignore svg:v-alphabetic 19.570,
    // ignore svg:v-hanging 19.571,
    // ignore svg:v-ideographic 19.572,
    // ignore svg:v-mathematical 19.573,
    // ignore svg:widths 19.576
    // ignore svg:x-height 19.580.
    attr: AttrMap2,
}

impl FontFaceDecl {
    /// New, empty.
    pub(crate) fn new_empty() -> Self {
        Self {
            name: "".to_string(),
            origin: Default::default(),
            attr: Default::default(),
        }
    }

    /// New, with a name.
    pub fn new<S: Into<String>>(name: S) -> Self {
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
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// The style:font-family-generic attribute specifies a generic font family name.
    /// The defined values for the style:font-family-generic attribute are:
    /// * decorative: the family of decorative fonts.
    /// * modern: the family of modern fonts.
    /// * roman: the family roman fonts (with serifs).
    /// * script: the family of script fonts.
    /// * swiss: the family roman fonts (without serifs).
    /// * system: the family system fonts.
    pub fn set_font_family_generic(&mut self, font: FontFamilyGeneric) {
        self.attrmap_mut()
            .set_attr("style:font-family-generic", font.to_string());
    }

    /// The style:font-pitch attribute specifies whether a font has a fixed or variable width.
    /// The defined values for the style:font-pitch attribute are:
    /// * fixed: font has a fixed width.
    /// * variable: font has a variable width.
    pub fn set_font_pitch(&mut self, pitch: FontPitch) {
        self.attrmap_mut()
            .set_attr("style:font-pitch", pitch.to_string());
    }

    /// External font family name.
    pub fn set_font_family<S: Into<String>>(&mut self, name: S) {
        self.attrmap_mut().set_attr("svg:font-family", name.into());
    }

    /// External font stretch value.
    pub fn set_font_stretch(&mut self, stretch: FontStretch) {
        self.attrmap_mut()
            .set_attr("svg:font-stretch", stretch.to_string());
    }

    /// External font style value.
    pub fn set_font_style(&mut self, style: FontStyle) {
        self.attrmap_mut()
            .set_attr("svg:font-style", style.to_string());
    }

    /// External font variant.
    pub fn set_font_variant(&mut self, variant: FontVariant) {
        self.attrmap_mut()
            .set_attr("svg:font-variant", variant.to_string());
    }

    /// External font weight.
    pub fn set_font_weight(&mut self, weight: FontWeight) {
        self.attrmap_mut()
            .set_attr("svg:font-weight", weight.to_string());
    }
}
