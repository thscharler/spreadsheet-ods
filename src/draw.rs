use crate::attrmap2::AttrMap2;
use crate::style::units::RelativeScale;
use crate::{CellRef, GraphicStyleRef, Length, ParagraphStyleRef};

/// The <office:annotation> element specifies an OpenDocument annotation. The annotation's
/// text is contained in <text:p> and <text:list> elements.
#[derive(Debug, Clone)]
pub struct Annotation {
    ///
    name: String,
    ///
    display: bool,
    ///
    attr: AttrMap2,
}

impl Annotation {
    pub fn new_empty() -> Self {
        Self {
            name: "".to_string(),
            display: true,
            attr: Default::default(),
        }
    }

    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            display: true,
            attr: Default::default(),
        }
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    /// Display
    pub fn display(&self) -> bool {
        self.display
    }

    /// Name
    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }

    draw_caption_point_x!(attr);
    draw_caption_point_y!(attr);
    draw_class_names!(attr);
    draw_corner_radius!(attr);
    draw_id!(attr);
    draw_layer!(attr);
    draw_style_name!(attr);
    draw_text_style_name!(attr);
    draw_transform!(attr);
    draw_z_index!(attr);
    svg_height!(attr);
    svg_width!(attr);
    svg_x!(attr);
    svg_y!(attr);
    table_end_cell_address!(attr);
    table_end_x!(attr);
    table_end_y!(attr);
    table_table_background!(attr);
    xml_id!(attr);
}

/// The <draw:rect> element represents a rectangular drawing shape.
#[derive(Debug, Clone)]
pub struct DrawRect {
    ///
    name: String,
    ///
    attr: AttrMap2,
}

impl DrawRect {
    pub fn new_empty() -> Self {
        Self {
            name: "".to_string(),
            attr: Default::default(),
        }
    }

    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            attr: Default::default(),
        }
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    draw_caption_id!(attr);
    draw_class_names!(attr);
    draw_corner_radius!(attr);
    draw_id!(attr);
    draw_layer!(attr);
    draw_style_name!(attr);
    draw_text_style_name!(attr);
    draw_transform!(attr);
    draw_z_index!(attr);
    svg_height!(attr);
    svg_width!(attr);
    svg_rx!(attr);
    svg_ry!(attr);
    svg_x!(attr);
    svg_y!(attr);
    table_end_cell_address!(attr);
    table_end_x!(attr);
    table_end_y!(attr);
    table_table_background!(attr);
    xml_id!(attr);
}

/// The <draw:frame> element represents a frame and serves as the container for elements that
/// may occur in a frame.
/// Frame formatting properties are stored in styles belonging to the graphic family.
#[derive(Debug, Clone)]
pub struct DrawFrame {
    ///
    name: String,
    ///
    attr: AttrMap2,
}

impl DrawFrame {
    pub fn new_empty() -> Self {
        Self {
            name: "".to_string(),
            attr: Default::default(),
        }
    }

    pub fn new<S: Into<String>>(name: S) -> Self {
        Self {
            name: name.into(),
            attr: Default::default(),
        }
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// Allows access to all attributes of the style itself.
    pub fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }

    /// Name
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Name
    pub fn set_name<S: Into<String>>(&mut self, name: S) {
        self.name = name.into();
    }

    draw_caption_id!(attr);
    draw_class_names!(attr);
    draw_corner_radius!(attr);
    draw_copy_of!(attr);
    draw_id!(attr);
    draw_layer!(attr);
    draw_style_name!(attr);
    draw_text_style_name!(attr);
    draw_transform!(attr);
    draw_z_index!(attr);
    style_rel_height!(attr);
    style_rel_width!(attr);
    svg_height!(attr);
    svg_width!(attr);
    svg_x!(attr);
    svg_y!(attr);
    table_end_cell_address!(attr);
    table_end_x!(attr);
    table_end_y!(attr);
    table_table_background!(attr);
    xml_id!(attr);
}
