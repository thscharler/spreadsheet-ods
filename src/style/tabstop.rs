//!
//! Defines tabstops for paragraph styles.
//!

use crate::attrmap2::AttrMap2;
use crate::style::color_string;
use crate::style::units::{Length, LineStyle, LineType, LineWidth, TabStopType};
use crate::TextStyleRef;
use color::Rgb;

/// The <style:tab-stops> element is a container for <style:tab-stop> elements.
/// If a style contains a <style:tab-stops> element, it overrides the entire <style:tab-stops>
/// element of the parent style such that no <style:tab-stop> children are inherited; otherwise,
/// the style inherits the entire <style:tab-stops> element as specified in section 16.2
/// <style:style>.
#[derive(Clone, Debug, Default)]
pub struct TabStop {
    attr: AttrMap2,
}

impl TabStop {
    /// Empty.
    pub fn new() -> Self {
        Self {
            attr: Default::default(),
        }
    }

    /// Delimiter character for tabs of type Char.
    pub fn set_tabstop_char(&mut self, c: char) {
        self.attr.set_attr("style:char", c.to_string());
    }

    /// The style:leader-color attribute specifies the color of a leader line. The value of this
    /// attribute is either font-color or a color. If the value is font-color, the current text color is
    /// used for the leader line.
    pub fn set_leader_color(&mut self, color: Rgb<u8>) {
        self.attr
            .set_attr("style:leader-color", color_string(color));
    }

    /// The style:leader-style attribute specifies a style for a leader line.
    ///
    /// The defined values for the style:leader-style attribute are:
    /// * none: tab stop has no leader line.
    /// * dash: tab stop has a dashed leader line.
    /// * dot-dash: tab stop has a leader line whose repeating pattern is a dot followed by a dash.
    /// * dot-dot-dash: tab stop has a leader line whose repeating pattern has two dots followed by
    /// a dash.
    /// * dotted: tab stop has a dotted leader line.
    /// * long-dash: tab stop has a dashed leader line whose dashes are longer than the ones from
    /// the dashed line for value dash.
    /// * solid: tab stop has a solid leader line.
    /// * wave: tab stop has a wavy leader line.
    ///
    /// Note: The definitions of the values of the style:leader-style attribute are based on the text
    /// decoration style 'text-underline-style' from [CSS3Text], §9.2.
    pub fn set_leader_style(&mut self, style: LineStyle) {
        self.attr.set_attr("style:leader-style", style.to_string());
    }

    /// The style:leader-text attribute specifies a single Unicode character for use as leader text
    /// for tab stops.
    /// An consumer may support only specific characters as textual leaders. If a character that is not
    /// supported by a consumer is specified by this attribute, the consumer should display a leader
    /// character that it supports instead of the one specified by this attribute.
    /// If both style:leader-text and style:leader-style 19.480 attributes are specified, the
    /// value of the style:leader-text sets the leader text for tab stops.
    ///
    /// The default value for this attribute is “ ” (U+0020, SPACE).
    pub fn set_leader_text(&mut self, text: char) {
        self.attr.set_attr("style:leader-text", text.to_string());
    }

    /// The style:leader-text-style specifies a text style that is applied to a textual leader. It is
    /// not applied to leader lines. If the attribute appears in an automatic style, it may reference either an
    /// automatic text style or a common style. If the attribute appears in a common style, it may
    /// reference a common style only.
    pub fn set_leader_text_style(&mut self, styleref: &TextStyleRef) {
        self.attr
            .set_attr("style:leader-text-style", styleref.to_string());
    }

    /// The style:leader-type attribute specifies whether a leader line should be drawn, and if so,
    /// whether a single or double line will be used.
    ///
    /// The defined values for the style:leader-type attribute are:
    /// * double: a double line is drawn.
    /// * none: no line is drawn.
    /// * single: a single line is drawn.
    pub fn set_leader_type(&mut self, t: LineType) {
        self.attr.set_attr("style:leader-type", t.to_string());
    }

    /// The style:leader-width attribute specifies the width (i.e., thickness) of a leader line.
    /// The defined values for the style:leader-width attribute are:
    /// * auto: the width of a leader line should be calculated from the font size of the text where the
    /// leader line will appear.
    /// * bold: the width of a leader line should be calculated from the font size of the text where the
    /// leader line will appear but is wider than for the value of auto.
    /// * a value of type percent 18.3.23
    /// * a value of type positiveInteger 18.2
    /// * a value of type positiveLength 18.3.26
    /// The line widths referenced by the values medium, normal, thick and thin are implementation defined.
    pub fn set_leader_width(&mut self, w: LineWidth) {
        self.attr.set_attr("style:leader-width", w.to_string());
    }

    /// The style:position attribute specifies the position of a tab stop. Depending on the value of
    /// the text:relative-tab-stop-position 19.861 attribute in the
    /// <text:table-ofcontent-source> 8.3.2,
    /// <text:illustration-index-source> 8.4.2,
    /// <text:object-index-source> 8.6.2,
    /// <text:user-index-source> 8.7.2 or
    /// <text:alphabetical-index-source> 8.8.2
    ///
    /// parent element, the position of the tab is interpreted as being relative to the left
    /// margin or the left indent.
    pub fn set_position(&mut self, pos: Length) {
        self.attr.set_attr("style:position", pos.to_string());
    }

    /// The style:type attribute specifies the type of a tab stop within paragraph formatting properties.
    /// The defined values for the style:type attribute are:
    /// * center: text is centered on a tab stop.
    /// * char: character appears at a tab stop position.
    /// * left: text is left aligned with a tab stop.
    /// * right: text is right aligned with a tab stop.
    /// For a <style:tab-stop> 17.8 element the default value for this attribute is left.
    pub fn set_tabstop_type(&mut self, t: TabStopType) {
        self.attr.set_attr("style:type", t.to_string());
    }

    /// General attributes.
    pub(crate) fn attrmap(&self) -> &AttrMap2 {
        &self.attr
    }

    /// General attributes.
    pub(crate) fn attrmap_mut(&mut self) -> &mut AttrMap2 {
        &mut self.attr
    }
}
