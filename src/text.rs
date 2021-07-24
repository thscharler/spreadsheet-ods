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
    pub fn outlinelevel(mut self, l: u8) -> Self {
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

text_tag!(TextSpan, "text:span");

impl TextSpan {
    /// Sets class names aka paragraph styles as formatting.
    pub fn class_names(mut self, class_names: &[&TextStyleRef]) -> Self {
        let mut buf = String::new();
        for n in class_names {
            buf.push_str(n.as_str());
            buf.push(' ');
        }
        self.xml.set_attr("text:class-names", buf);
        self
    }

    /// Style for this paragraph.
    pub fn style_name(mut self, name: &TextStyleRef) -> Self {
        self.xml.set_attr("text:style-name", name.to_string());
        self
    }
}

text_tag!(TextA, "text:a");

impl TextA {
    /// href for a link.
    pub fn href<S: Into<String>>(mut self, uri: S) -> Self {
        self.xml.set_attr("xlink:href", uri.into());
        self
    }

    /// Style name.
    pub fn style_name(mut self, style: &TextStyleRef) -> Self {
        self.xml.set_attr("text:style-name", style.to_string());
        self
    }

    /// Style name for a visited link.
    pub fn visited_style_name(mut self, style: &TextStyleRef) -> Self {
        self.xml
            .set_attr("text:visited-style-name", style.to_string());
        self
    }
}

text_tag!(TextS, "text:s");

impl TextS {
    /// Number of spaces.
    pub fn count(mut self, count: u32) -> Self {
        self.xml.set_attr("text:c", count.to_string());
        self
    }
}

text_tag!(TextTab, "text:tab");

impl TextTab {
    /// To which tabstop does this refer to?
    pub fn tab_ref(mut self, tab_ref: u32) -> Self {
        self.xml.set_attr("text:tab-ref", tab_ref.to_string());
        self
    }
}

text_tag!(TextLineBreak, "text:line-break");
text_tag!(SoftPageBreak, "text:soft-page-break");

text_tag!(AuthorInitials, "text:author-initials");
text_tag!(AuthorName, "text:author_name");
text_tag!(Chapter, "text:chapter");
text_tag!(CharacterCount, "text:character-count");
text_tag!(CreationDate, "text:creation-date");
text_tag!(CreationTime, "text:creation-time");
text_tag!(Creator, "text:creator");
text_tag!(Date, "text:date");
text_tag!(Description, "text:description");
text_tag!(EditingCycles, "text:editing-cycles");
text_tag!(EditingDuration, "text:editingduration");
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

// <office:annotation> 14.1,
//  <office:annotation-end> 14.2,
//  <text:alphabetical-index-mark> 8.1.10,
//  <text:alphabeticalindex-mark-end> 8.1.9,
//  <text:alphabetical-index-mark-start> 8.1.8,
// <text:bibliography-mark> 8.1.11,
//  <text:bookmark> 6.2.1.2,
//  <text:bookmark-end> 6.2.1.4,
//  <text:bookmark-ref> 7.7.6,
//  <text:bookmark-start> 6.2.1.3,
//  <text:change> 5.5.8.4,
//  <text:change-end> 5.5.8.3,
//  <text:change-start> 5.5.8.2,
//  <text:conditional-text> 7.7.3,
// <text:database-display> 7.6.3,
//  <text:database-name> 7.6.7,
//  <text:databasenext> 7.6.4,
//  <text:database-row-number> 7.6.6,
//  <text:database-row-select> 7.6.5,
//  <text:dde-connection> 7.7.12,
// <text:drop-down> 7.4.16,
//  <text:execute-macro> 7.7.10,
//  <text:expression> 7.4.14,
//  <text:hidden-paragraph> 7.7.11,
//  <text:hidden-text> 7.7.4,
//  <text:image-count> 7.5.18.7,
//  <text:measure> 7.7.13,
// <text:meta> 6.1.9,
//  <text:meta-field> 7.5.19,
//  <text:note> 6.3.2,
//  <text:note-ref> 7.7.7,
// <text:object-count> 7.5.18.8,
//  <text:page-continuation> 7.3.5,
//  <text:page-variable-get> 7.7.1.3,
// <text:page-variable-set> 7.7.1.2,
//  <text:paragraph-count> 7.5.18.3,
// <text:placeholder> 7.7.2,
//  <text:reference-mark> 6.2.2.2,
//  <text:reference-markend> 6.2.2.4,
//  <text:reference-mark-start> 6.2.2.3,
//  <text:reference-ref> 7.7.5,
// <text:ruby> 6.4,
//  <text:script> 7.7.9,
//  <text:sender-city> 7.3.6.13,
// <text:sender-company> 7.3.6.10,
//  <text:sender-country> 7.3.6.15,
//  <text:senderemail> 7.3.6.7,
//  <text:sender-fax> 7.3.6.9,
//  <text:sender-firstname> 7.3.6.2,
// <text:sender-initials> 7.3.6.4,
//  <text:sender-lastname> 7.3.6.3,
//  <text:senderphone-private> 7.3.6.8,
//  <text:sender-phone-work> 7.3.6.11,
//  <text:senderposition> 7.3.6.6,
//  <text:sender-postal-code> 7.3.6.14,
//  <text:sender-state-orprovince> 7.3.6.16,
//  <text:sender-street> 7.3.6.12,
//  <text:sender-title> 7.3.6.5,
// <text:sequence> 7.4.13,
//  <text:sequence-ref> 7.7.8,
//  <text:table-formula> 7.7.14,
// <text:template-name> 7.3.10,
//  <text:text-input> 7.4.15,
//  <text:toc-mark> 8.1.4,
//  <text:toc-mark-end> 8.1.3,
// <text:toc-mark-start> 8.1.2,
//  <text:user-defined> 7.5.6,
//  <text:user-field-get> 7.4.9,
//  <text:user-field-input> 7.4.10,
//  <text:user-index-mark> 8.1.7,
// <text:user-index-mark-end> 8.1.6,
//  <text:user-index-mark-start> 8.1.5,
// <text:variable-get> 7.4.5,
//  <text:variable-input> 7.4.6,
//  <text:variable-set> 7.4.4 and
