use crate::attrmap::{AttrMap, AttrMapIter, AttrMapType};

/// A tag within a text region.
#[derive(Debug, Clone, Default)]
pub struct TextTag {
    tag: String,
    attr: Option<AttrMapType>,
}

impl AttrMap for TextTag {
    fn attr_map(&self) -> Option<&AttrMapType> {
        self.attr.as_ref()
    }

    fn attr_map_mut(&mut self) -> &mut Option<AttrMapType> {
        &mut self.attr
    }
}

impl TextTag {
    pub fn new<S: Into<String>>(tag: S) -> Self {
        Self {
            tag: tag.into(),
            attr: None,
        }
    }

    pub fn set_tag<S: Into<String>>(&mut self, tag: S) {
        self.tag = tag.into();
    }

    pub fn tag(&self) -> &String {
        &self.tag
    }

    pub fn attr_iter(&self) -> AttrMapIter {
        AttrMapIter::from(self.attr_map())
    }
}

/// Complex text is laid out as a sequence of tags, end-tags and text.
/// The user of this must ensure that the result is valid xml.
#[derive(Debug, Clone)]
pub enum TextElem {
    Start(TextTag),
    Empty(TextTag),
    Text(String),
    End(String),
}

/// A vector of text.
#[derive(Debug, Clone, Default)]
pub struct TextVec {
    vec: Option<Vec<TextElem>>,
}

impl TextVec {
    /// Create.
    pub fn new() -> Self {
        Self {
            vec: None
        }
    }

    /// Append to the vector
    pub fn start<S: Into<String>>(&mut self, ctag: S) {
        self.push(TextElem::Start(TextTag::new(ctag.into())));
    }

    /// Append to the vector
    pub fn startc(&mut self, ctag: TextTag) {
        self.push(TextElem::Start(ctag));
    }

    /// Append to the vector
    pub fn empty<S: Into<String>>(&mut self, ctag: S) {
        self.push(TextElem::Empty(TextTag::new(ctag.into())));
    }

    /// Append to the vector
    pub fn emptyc(&mut self, ctag: TextTag) {
        self.push(TextElem::Empty(ctag));
    }

    /// Append to the vector
    pub fn text<S: Into<String>>(&mut self, txt: S) {
        self.push(TextElem::Text(txt.into()));
    }

    /// Append to the vector
    pub fn end<S: Into<String>>(&mut self, ctag: S) {
        self.push(TextElem::End(ctag.into()));
    }

    /// Append to the vector
    pub fn push(&mut self, cm: TextElem) {
        self.vec
            .get_or_insert_with(Vec::new)
            .push(cm);
    }

    /// Remove all content.
    pub fn clear(&mut self) {
        self.vec = None;
    }

    /// No vec contained.
    pub fn is_empty(&self) -> bool {
        self.vec.is_none()
    }

    /// Checks if this is a valid sequence of text, in way that it
    /// can be written to output without destroying the xml.
    pub fn is_valid(&self, open_tag: &mut String, close_tag: &mut String) -> bool {
        let mut res = true;

        let mut tags = Vec::new();

        if let Some(vec) = &self.vec {
            for c in vec {
                match c {
                    TextElem::Start(t) =>
                        tags.push(t.tag.clone()),
                    TextElem::End(t) => {
                        let tag = tags.pop();
                        if let Some(ref tag) = tag {
                            if t != tag {
                                std::mem::swap(open_tag, &mut tag.clone());
                                std::mem::swap(close_tag, &mut t.clone());
                                res = false;
                                break;
                            }
                        } else {
                            res = false;
                            break;
                        }
                    }
                    _ => (),
                }
            }
        }

        res
    }

    /// Returns the text vec itself.
    pub fn vec(&self) -> Option<&Vec<TextElem>> {
        self.vec.as_ref()
    }
}