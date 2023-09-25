//! Document metadata.

use crate::xlink::{XLinkActuate, XLinkShow, XLinkType};
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

/// Metadata
#[derive(Debug, Clone)]
pub struct Metadata {
    /// The <meta:generator> element contains a string that identifies the OpenDocument producer
    /// that was used to create or last modify the document. This string should match the definition for
    /// user-agents in the HTTP protocol as specified in section 14.43 of RFC2616. The generator string
    /// should allow OpenDocument consumers to distinguish between all released versions of a
    /// producer.
    /// Note: Release versions of a producer could be distinguished based on build
    /// identifiers or patch-level information.
    /// If an OpenDocument producer that creates a document cannot provide an identifier string, the
    /// producer shall not export this element. If a producer stores a modified document created by
    /// another producer cannot provide a unique identifier, it shall not export the original identifier
    /// belonging to the producer that created the document.
    pub generator: String,
    /// The dc:title element specifies the title of a document
    pub title: String,
    /// The dc:description element contains a description of a document.
    pub description: String,
    /// The dc:subject element specifies the subject of a document
    pub subject: String,
    /// The meta:keyword element contains a keyword pertaining to a document.
    pub keyword: String,
    /// The meta:initial-creator element specifies the name of the initial creator of a document
    pub initial_creator: String,
    /// The dc:creator element specifies the name of the person who last modified a
    /// document, who created an annotation, who authored a change .
    pub creator: String,
    /// The meta:printed-by element specifies the name of the last person who printed a
    /// document.
    pub printed_by: String,
    /// The meta:creation-date element specifies the date and time when a document was
    /// created.
    pub creation_date: Option<NaiveDateTime>,
    /// The dc:date element specifies the date and time when the document was last
    /// modified, when an annotation was created, when a change was made.
    pub date: Option<NaiveDateTime>,
    /// The meta:print-date element specifies the date and time when a document was last
    /// printed.
    pub print_date: Option<NaiveDateTime>,
    /// The dc:language element specifies the default language of a document
    pub language: String,
    /// The meta:editing-cycles element specifies the number of times a document has been
    /// edited. When a document is created, this value is set to 1. Each time a document is saved, the
    /// editing-cycles number is incremented by 1.
    pub editing_cycles: u32,
    /// The meta:editing-duration element specifies the total time spent editing a document.
    pub editing_duration: Duration,
    /// The <meta:template> element specifies an IRI for the document template that was used to
    /// create a document. The IRI is specified using the xlink:href attribute.
    pub template: MetaTemplate,
    /// The meta:auto-reload element specifies whether a document is reloaded or replaced by
    /// another document after a specified period of time has elapsed.
    pub auto_reload: MetaAutoReload,
    /// The meta:hyperlink-behaviour element specifies the default behavior for hyperlinks in a
    /// document.
    pub hyperlink_behaviour: MetaHyperlinkBehaviour,
    /// The meta:document-statistic element represents statistics about a document.
    pub document_statistics: MetaDocumentStatistics,
    /// The <meta:user-defined> element specifies any additional user-defined metadata for a
    /// document.
    pub user_defined: Vec<MetaUserDefined>,
}

impl Default for Metadata {
    fn default() -> Self {
        Self {
            generator: Default::default(),
            title: Default::default(),
            description: Default::default(),
            subject: Default::default(),
            keyword: Default::default(),
            initial_creator: Default::default(),
            creator: Default::default(),
            printed_by: Default::default(),
            creation_date: None,
            date: None,
            print_date: None,
            template: Default::default(),
            auto_reload: Default::default(),
            hyperlink_behaviour: Default::default(),
            document_statistics: Default::default(),
            language: Default::default(),
            editing_cycles: 0,
            editing_duration: Duration::seconds(0),
            user_defined: vec![],
        }
    }
}

/// Specifies an IRI for the document template that was used to
/// create a document.
#[derive(Debug, Default, Clone)]
pub struct MetaTemplate {
    /// The meta:date attribute specifies the date and time when a template was last modified, prior to
    /// being used to create the current document.
    pub date: Option<NaiveDateTime>,
    /// See XLink
    pub actuate: Option<XLinkActuate>,
    /// See XLink
    pub href: Option<String>,
    /// See XLink
    pub title: Option<String>,
    /// See XLink
    pub link_type: Option<XLinkType>,
}

impl MetaTemplate {
    /// Everything is None.
    pub fn is_empty(&self) -> bool {
        self.date.is_none()
            && self.actuate.is_none()
            && self.href.is_none()
            && self.title.is_none()
            && self.link_type.is_none()
    }
}

/// Specifies whether a document is reloaded or replaced by
/// another document after a specified period of time has elapsed.
#[derive(Debug, Default, Clone)]
pub struct MetaAutoReload {
    /// The meta:delay attribute specifies a reload delay.
    pub delay: Option<Duration>,
    /// See XLink
    pub actuate: Option<XLinkActuate>,
    /// See XLink
    pub href: Option<String>,
    /// See XLink
    pub show: Option<XLinkShow>,
    /// See XLink
    pub link_type: Option<XLinkType>,
}

impl MetaAutoReload {
    /// Everything is None.
    pub fn is_empty(&self) -> bool {
        self.delay.is_none()
            && self.actuate.is_none()
            && self.href.is_none()
            && self.show.is_none()
            && self.link_type.is_none()
    }
}

/// Specifies the default behavior for hyperlinks in a document.
#[derive(Debug, Default, Clone)]
pub struct MetaHyperlinkBehaviour {
    /// The office:target-frame-name attribute specifies the name of a target frame.
    /// The defined values for the office:target-frame-name attribute are:
    /// • _blank: The referenced document is displayed in a new frame.
    /// • _parent: The referenced document is displayed in the parent frame of the current frame.
    /// • _self: The referenced document replaces the content of the current frame.
    /// • _top: The referenced document is displayed in the topmost frame, that is the frame that
    /// contains the current frame as a child or descendant but is not contained within another frame.
    /// • A frame name: The referenced document is displayed in the named frame. If the named frame
    /// does not exist, a new frame with that name is created.
    /// The office:target-frame-name attribute may be used together with an xlink:show 19.917
    /// attribute. In that case, if the value of the attribute is _blank, the xlink:show attribute value
    /// should be new. If the value of the attribute is any of the other value options, the value of the
    /// xlink:show attribute should be replace.
    pub target_frame_name: Option<String>,
    /// See XLink
    pub show: Option<XLinkShow>,
}

impl MetaHyperlinkBehaviour {
    /// Everything is None.
    pub fn is_empty(&self) -> bool {
        self.target_frame_name.is_none() && self.show.is_none()
    }
}

/// Represents statistics about a document.
#[derive(Debug, Default, Clone)]
pub struct MetaDocumentStatistics {
    ///
    pub cell_count: u32,
    ///
    pub object_count: u32,
    ///
    pub ole_object_count: u32,
    ///
    pub table_count: u32,
}

/// Specifies any additional user-defined metadata for a document.
#[derive(Debug, Clone)]
pub struct MetaUserDefined {
    /// Name
    pub name: String,
    /// Value
    pub value: MetaValue,
}

impl Default for MetaUserDefined {
    fn default() -> Self {
        Self {
            name: "".to_string(),
            value: MetaValue::String("".to_string()),
        }
    }
}

/// Value for user defined metadata.
#[derive(Debug, Clone)]
pub enum MetaValue {
    ///
    Boolean(bool),
    ///
    Date(NaiveDate),
    ///
    Float(f64),
    ///
    Time(NaiveTime),
    ///
    String(String),
}
