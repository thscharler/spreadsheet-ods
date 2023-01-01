//! Text is stored as a simple String whenever possible.
//! When there is a more complex structure, a TextTag is constructed
//! which mirrors the Xml tree structure.
//!
//! For construction of a new TextTag structure a few helper structs are
//! defined.
//!
//! ```
//! use spreadsheet_ods::text::{TextP, TextTag, AuthorName, CreationDate, TextS};
//! use spreadsheet_ods::style::ParagraphStyleRef;
//!
//! let p1_ref = ParagraphStyleRef::from("p1");
//!
//! let txt = TextP::new()
//!             .style_name(&p1_ref)
//!             .text("some text")
//!             .tag(AuthorName::new())
//!             .tag(TextS::new())
//!             .tag(CreationDate::new())
//!             .tag(TextS::new())
//!             .text("whatever");
//! println!("{}", txt.into_xmltag());
//! ```
//!

use crate::style::{ParagraphStyleRef, TextStyleRef};
use crate::xmltree::{XmlContent, XmlTag};
use std::fmt::{Display, Formatter};

/// TextTags are just XmlTags.
pub type TextTag = XmlTag;
/// Content of a TextTag is just some XmlContent.
pub type TextContent = XmlContent;

text_tag!(TextH, "text:h");

// ok text:class-names 19.770.2,
// ok text:cond-style-name 19.776,
// ok text:id 19.809.6,
// ok text:is-list-header 19.816,
// ok text:outline-level 19.844.4,
// ok text:restart-numbering 19.857,
// ok text:start-value 19.868.2,
// ok text:style-name 19.874.7,
// ignore xhtml:about 19.905,
// ignore xhtml:content 19.906,
// ignore xhtml:datatype 19.907,
// ignore xhtml:property 19.908
// ok xml:id 19.914.
impl TextH {
    /// Sets class names aka paragraph styles as formatting.
    pub fn class_names(mut self, class_names: &[&ParagraphStyleRef]) -> Self {
        let mut buf = String::new();
        for n in class_names {
            buf.push_str(n.as_str());
            buf.push(' ');
        }
        self.xml.set_attr("text:class-names", buf);
        self
    }

    /// Sets a conditional style.
    pub fn condstyle_name(mut self, name: &ParagraphStyleRef) -> Self {
        self.xml.set_attr("text:condstyle-name", name.to_string());
        self
    }

    /// Identifier for a text passage.
    pub fn id(mut self, id: &str) -> Self {
        self.xml.set_attr("text:id", id);
        self
    }

    /// Styled as list header.
    pub fn list_header(mut self, lh: bool) -> Self {
        self.xml.set_attr("text:is-list-header", lh.to_string());
        self
    }

    /// Level of the heading.
    pub fn outline_level(mut self, l: u8) -> Self {
        self.xml.set_attr("text:outlinelevel", l.to_string());
        self
    }

    /// Numbering reset.
    pub fn restart_numbering(mut self, r: bool) -> Self {
        self.xml.set_attr("text:restart-numbering", r.to_string());
        self
    }

    /// Numbering start value.
    pub fn start_value(mut self, l: u8) -> Self {
        self.xml.set_attr("text:start-value", l.to_string());
        self
    }

    /// Style
    pub fn style_name(mut self, name: &ParagraphStyleRef) -> Self {
        self.xml.set_attr("text:style-name", name.to_string());
        self
    }

    /// xml-id
    pub fn xml_id(mut self, id: &str) -> Self {
        self.xml.set_attr("xml:id", id);
        self
    }
}

text_tag!(TextP, "text:p");

// ok text:class-names 19.770.3,
// ok text:cond-style-name 19.776,
// ok text:id 19.809.8,
// ok text:style-name 19.874.29,
// ignore xhtml:about 19.905,
// ignore xhtml:content 19.906,
// ignore xhtml:datatype 19.907,
// ignore xhtml:property 19.908
// ok xml:id 19.914.
impl TextP {
    /// Sets class names aka paragraph styles as formatting.
    pub fn class_names(mut self, class_names: &[&ParagraphStyleRef]) -> Self {
        let mut buf = String::new();
        for n in class_names {
            buf.push_str(n.as_str());
            buf.push(' ');
        }
        self.xml.set_attr("text:class-names", buf);
        self
    }

    /// Sets a conditional style.
    pub fn condstyle_name(mut self, name: &ParagraphStyleRef) -> Self {
        self.xml.set_attr("text:condstyle-name", name.to_string());
        self
    }

    /// Text id for a text passage.
    pub fn id(mut self, id: &str) -> Self {
        self.xml.set_attr("text:id", id);
        self
    }

    /// Style for this paragraph.
    pub fn style_name(mut self, name: &ParagraphStyleRef) -> Self {
        self.xml.set_attr("text:style-name", name.to_string());
        self
    }

    /// xml-id
    pub fn xml_id(mut self, id: &str) -> Self {
        self.xml.set_attr("xml:id", id);
        self
    }
}

// The <text:span> element represents the application of a style to the character data of a portion
// of text. The content of this element is the text which uses that text style.
//
// The <text:span> element can be nested.
//
// White space characters contained in this element are collapsed.
text_tag!(TextSpan, "text:span");

// text:class-names 19.770.4 and
// text:style-name 19.874.33.
impl TextSpan {
    /// A text:class-names attribute specifies a white space separated list of text style names.
    pub fn class_names(mut self, class_names: &[&TextStyleRef]) -> Self {
        let mut buf = String::new();
        for n in class_names {
            buf.push_str(n.as_str());
            buf.push(' ');
        }
        self.xml.set_attr("text:class-names", buf);
        self
    }

    /// The text:style-name attribute specifies style for span which shall be a style with family of
    /// text.
    /// If both text:style-name and text:class-names are present, the style referenced by the
    /// text:style-name attribute is treated as the first style in the list in text:class-names.
    /// Consumers should support the text:class-names attribute and also should preserve it while
    /// editing.
    pub fn style_name(mut self, name: &TextStyleRef) -> Self {
        self.xml.set_attr("text:style-name", name.to_string());
        self
    }
}

// The <text:a> element represents a hyperlink.
//
// The anchor of a hyperlink is composed of the character data contained by the <text:a> element
// and any of its descendant elements which constitute the character data of the paragraph which
// contains the <text:a> element. 6.1.1
text_tag!(TextA, "text:a");

// obsolete office:name 19.376.9,
// ??? office:target-frame-name 19.381,
// ??? office:title 19.383,
// ok text:style-name 19.874.2,
// ok text:visited-style-name 19.901,
// ??? xlink:actuate 19.909,
// ok xlink:href 19.910.33,
// ??? xlink:show 19.911 and
// ??? xlink:type 19.913.
impl TextA {
    /// The text:style-name attribute specifies a text style for an unvisited hyperlink.
    pub fn style_name(mut self, style: &TextStyleRef) -> Self {
        self.xml.set_attr("text:style-name", style.to_string());
        self
    }

    /// The text:visited-style-name attribute specifies a style for a hyperlink that has been visited.
    pub fn visited_style_name(mut self, style: &TextStyleRef) -> Self {
        self.xml
            .set_attr("text:visited-style-name", style.to_string());
        self
    }

    /// href for a link.
    pub fn href<S: Into<String>>(mut self, uri: S) -> Self {
        self.xml.set_attr("xlink:href", uri.into());
        self
    }
}

// The <text:s> element is used to represent the [UNICODE] character “ “ (U+0020, SPACE).
// This element shall be used to represent the second and all following “ “ (U+0020, SPACE)
// characters in a sequence of “ “ (U+0020, SPACE) characters.
//
// Note: It is not an error if the character preceding the element is not a white space character, but it
// is good practice to use this element only for the second and all following “ “ (U+0020, SPACE)
// characters in a sequence.
text_tag!(TextS, "text:s");

// text:c 19.763.
impl TextS {
    /// The text:c attribute specifies the number of “ “ (U+0020, SPACE) characters that a <text:s>
    /// element represents. A missing text:c attribute is interpreted as a single “ “ (U+0020, SPACE)
    /// character.
    pub fn count(mut self, count: u32) -> Self {
        self.xml.set_attr("text:c", count.to_string());
        self
    }
}

// The <text:tab> element represents the [UNICODE] tab character (HORIZONTAL
// TABULATION, U+0009).
//
// A <text:tab> element specifies that content immediately following it
// should begin at the next tab stop.
text_tag!(TextTab, "text:tab");

impl TextTab {
    /// The text:tab-ref attribute contains the number of the tab-stop to which a tab character refers.
    /// The position 0 marks the start margin of a paragraph.
    ///
    /// Note: The text:tab-ref attribute is only a hint to help non-layout oriented consumers to
    /// determine the tab/tab-stop association. Layout oriented consumers should determine the tab
    /// positions based on the style information.
    pub fn tab_ref(mut self, tab_ref: u32) -> Self {
        self.xml.set_attr("text:tab-ref", tab_ref.to_string());
        self
    }
}

// TODO: more of this.

// The <text:line-break> element represents a line break
text_tag!(TextLineBreak, "text:line-break");
// The <text:soft-page-break> element represents a soft page break within or between
// paragraph elements. As a child element of a <table:table> element it represents a soft page break between two
// table rows. It may appear in front of a <table:table-row> element.
text_tag!(SoftPageBreak, "text:soft-page-break");
// The <text:author-initials> element represents the initials of the author of a document.
text_tag!(AuthorInitials, "text:author-initials");
// The <text:author-name> element represents the full name of the author of a document.
text_tag!(AuthorName, "text:author_name");
//
text_tag!(Chapter, "text:chapter");
text_tag!(CharacterCount, "text:character-count");
text_tag!(CreationDate, "text:creation-date");
text_tag!(CreationTime, "text:creation-time");
text_tag!(Creator, "text:creator");
// The <text:date> element displays a date, by default this is the current date. The date can be
// adjusted to display a date other than the current date.
text_tag!(Date, "text:date");
text_tag!(Description, "text:description");
text_tag!(EditingCycles, "text:editing-cycles");
text_tag!(EditingDuration, "text:editing-duration");
text_tag!(FileName, "text:file-name");
text_tag!(InitialCreator, "text:initial-creator");
text_tag!(Keywords, "text:keywords");
text_tag!(ModificationDate, "text:modification-date");
text_tag!(ModificationTime, "text:modification-time");
text_tag!(PageCount, "text:pagecount");
text_tag!(PageNumber, "text:page-number");
text_tag!(PrintDate, "text:print-date");
text_tag!(PrintedBy, "text:printed-by");
text_tag!(PrintTime, "text:print-time");
text_tag!(SheetName, "text:sheet-name");
text_tag!(Subject, "text:subject");
text_tag!(TableCount, "text:table-count");
text_tag!(Time, "text:time");
text_tag!(Title, "text:title");

// Child elements of text:p
//
// <office:annotation> 14.1
// <office:annotation-end> 14.2
// <presentation:datetime> 10.9.3.5
// <presentation:footer> 10.9.3.3
// <presentation:header> 10.9.3.1,
// ok <text:a> 6.1.8
// <text:alphabetical-index-mark> 8.1.10
// <text:alphabeticalindex-mark-end> 8.1.9
// <text:alphabetical-index-mark-start> 8.1.8,
// ok <text:author-initials> 7.3.7.2
// ok <text:author-name> 7.3.7.1,
// <text:bibliography-mark> 8.1.11
// <text:bookmark> 6.2.1.2
// <text:bookmark-end> 6.2.1.4
// <text:bookmark-ref> 7.7.6
// <text:bookmark-start> 6.2.1.3
// <text:change> 5.5.7.4
// <text:change-end> 5.5.7.3
// <text:change-start> 5.5.7.2
// ok <text:chapter> 7.3.8
// ok <text:character-count> 7.5.18.5
// <text:conditional-text> 7.7.3,
// ok <text:creation-date> 7.5.3
// ok <text:creation-time> 7.5.4
// ok <text:creator> 7.5.17,
// <text:database-display> 7.6.3
// <text:database-name> 7.6.7
// <text:databasenext> 7.6.4
// <text:database-row-number> 7.6.6
// <text:database-row-select> 7.6.5
// ok <text:date> 7.3.2
// <text:dde-connection> 7.7.12
// ok <text:description> 7.5.5,
// ok <text:editing-cycles> 7.5.13
// ok <text:editing-duration> 7.5.14
// <text:executemacro> 7.7.10
// <text:expression> 7.4.14
// ok <text:file-name> 7.3.9
// <text:hiddenparagraph> 7.7.11
// <text:hidden-text> 7.7.4
// <text:image-count> 7.5.18.7,
// ok <text:initial-creator> 7.5.2
// ok <text:keywords> 7.5.12
// ok <text:line-break> 6.1.5,
// <text:measure> 7.7.13
// <text:meta> 6.1.9
// <text:meta-field> 7.5.19,
// ok <text:modification-date> 7.5.16
// ok <text:modification-time> 7.5.15
// <text:not_e> 6.3.2
// <text:not_e-ref> 7.7.7
// <text:object-count> 7.5.18.8
// <text:pagecontinuation> 7.3.5
// ok <text:page-count> 7.5.18.2
// ok <text:page-number> 7.3.4,
// <text:page-variable-get> 7.7.1.3
// <text:page-variable-set> 7.7.1.2,
// <text:paragraph-count> 7.5.18.3
// <text:placeholder> 7.7.2
// ok <text:print-date> 7.5.8
// ok <text:printed-by> 7.5.9
// ok <text:print-time> 7.5.7
// <text:reference-mark> 6.2.2.2
// <text:reference-mark-end> 6.2.2.4
// <text:reference-mark-start> 6.2.2.3,
// <text:reference-ref> 7.7.5
// <text:ruby> 6.4
// ok <text:s> 6.1.3
// <text:script> 7.7.9,
// <text:sender-city> 7.3.6.13
// <text:sender-company> 7.3.6.10
// <text:sendercountry> 7.3.6.15
// <text:sender-email> 7.3.6.7
// <text:sender-fax> 7.3.6.9,
// <text:sender-firstname> 7.3.6.2
// <text:sender-initials> 7.3.6.4
// <text:senderlastname> 7.3.6.3
// <text:sender-phone-private> 7.3.6.8
// <text:sender-phonework> 7.3.6.11
// <text:sender-position> 7.3.6.6
// <text:sender-postal-code> 7.3.6.14
// <text:sender-state-or-province> 7.3.6.16
// <text:sender-street> 7.3.6.12
// <text:sender-title> 7.3.6.5
// <text:sequence> 7.4.13
// <text:sequenceref> 7.7.8
// ok <text:sheet-name> 7.3.11
// <text:soft-page-break> 5.6
// ok <text:span> 6.1.7
// ok <text:subject> 7.5.11
// ok <text:tab> 6.1.4
// ok <text:table-count> 7.5.18.6,
// <text:table-formula> 7.7.14
// <text:template-name> 7.3.10
// <text:text-input> 7.4.15
// ok <text:time> 7.3.3
// ok <text:title> 7.5.10
// <text:toc-mark> 8.1.4
// <text:tocmark-end> 8.1.3
// <text:toc-mark-start> 8.1.2
// <text:user-defined> 7.5.6,
// <text:user-field-get> 7.4.9
// <text:user-field-input> 7.4.10
// <text:userindex-mark> 8.1.7
// <text:user-index-mark-end> 8.1.6
// <text:user-index-markstart> 8.1.5
// <text:variable-get> 7.4.5
// <text:variable-input> 7.4.6,
// <text:variable-set> 7.4.4
// <text:word-count> 7.5.18.4.
//
