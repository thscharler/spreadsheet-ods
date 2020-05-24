use std::io::{self, Write};
use std::fmt;
use std::marker::PhantomData;
use std::fmt::{Display, Formatter};

pub type Result = io::Result<()>;

#[derive(PartialEq)]
enum Open {
    None,
    Elem,
    Empty,
}

impl Display for Open {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        match self {
            Open::None => f.write_str("None")?,
            Open::Elem => f.write_str("Elem")?,
            Open::Empty => f.write_str("Empty")?,
        }
        Ok(())
    }
}

#[derive(Debug)]
struct Stack<'a> {
    #[cfg(feature = "check_xml")]
    stack: Vec<&'a str>,
    #[cfg(not(feature = "check_xml"))]
    stack: PhantomData<&'a str>,
}

#[cfg(feature = "check_xml")]
impl<'a> Stack<'a> {
    fn new() -> Self {
        Self {
            stack: Vec::new()
        }
    }

    fn len(&self) -> usize {
        self.stack.len()
    }

    fn push(&mut self, name: &'a str) {
        self.stack.push(name);
    }

    fn pop(&mut self) -> Option<&'a str> {
        self.stack.pop()
    }

    fn is_empty(&self) -> bool {
        self.stack.is_empty()
    }
}

#[cfg(not(feature = "check_xml"))]
impl<'a> Stack<'a> {
    fn new() -> Self {
        Self {
            stack: PhantomData {}
        }
    }

    fn len(&self) -> usize {
        0
    }

    fn push(&mut self, _name: &'a str) {}

    fn pop(&mut self) -> Option<&'a str> {
        None
    }

    fn is_empty(&self) -> bool {
        true
    }
}

/// The XmlWriter himself
pub struct XmlWriter<'a, W: Write> {
    writer: Box<W>,
    buf: String,
    stack: Stack<'a>,
    open: Open,
    /// if `true` it will indent all opening elements
    pub indent: bool,
}

impl<'a, W: Write> fmt::Debug for XmlWriter<'a, W> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        Ok(write!(f, "XmlWriter {{ stack: {:?}, opened: {} }}", self.stack, self.open)?)
    }
}

impl<'a, W: Write> XmlWriter<'a, W> {
    /// Create a new writer, by passing an `io::Write`
    pub fn new(writer: W) -> XmlWriter<'a, W> {
        XmlWriter {
            stack: Stack::new(),
            buf: String::new(),
            writer: Box::new(writer),
            open: Open::None,
            indent: true,
        }
    }

    /// Write the DTD. You have to take care of the encoding
    /// on the underlying Write yourself.
    pub fn dtd(&mut self, encoding: &str) -> Result {
        self.buf.push_str("<?xml version=\"1.0\" encoding=\"");
        self.buf.push_str(encoding);
        self.buf.push_str("\" ?>\n");

        Ok(())
    }

    fn indent(&mut self) {
        if cfg!(feature = "indent_xml") && self.indent && !self.stack.is_empty() {
            self.buf.push('\n');
            let indent = self.stack.len() * 2;
            for _ in 0..indent {
                self.buf.push(' ');
            };
        }
    }

    /// Write an element with inlined text (not escaped)
    pub fn elem_text(&mut self, name: &str, text: &str) -> Result {
        self.close_elem()?;

        self.indent();

        self.buf.push('<');
        self.buf.push_str(name);
        self.buf.push('>');

        self.buf.push_str(text);

        self.buf.push('<');
        self.buf.push('/');
        self.buf.push_str(name);
        self.buf.push('>');

        Ok(())
    }

    /// Write an element with inlined text (escaped)
    pub fn elem_text_esc(&mut self, name: &str, text: &str) -> Result {
        self.close_elem()?;

        self.indent();

        self.buf.push('<');
        self.buf.push_str(name);
        self.buf.push('>');

        self.escape(text, false);

        self.buf.push('<');
        self.buf.push('/');
        self.buf.push_str(name);
        self.buf.push('>');

        Ok(())
    }

    /// Begin an elem, make sure name contains only allowed chars
    pub fn elem(&mut self, name: &'a str) -> Result {
        self.close_elem()?;

        self.indent();

        self.stack.push(name);

        self.buf.push('<');
        self.open = Open::Elem;
        self.buf.push_str(name);
        Ok(())
    }

    /// Begin an empty elem
    pub fn empty(&mut self, name: &'a str) -> Result {
        self.close_elem()?;

        self.indent();

        self.buf.push('<');
        self.open = Open::Empty;
        self.buf.push_str(name);
        Ok(())
    }

    /// Close an elem if open, do nothing otherwise
    fn close_elem(&mut self) -> Result {
        match self.open {
            Open::None => {}
            Open::Elem => {
                self.buf.push('>');
            }
            Open::Empty => {
                self.buf.push('/');
                self.buf.push('>');
            }
        }
        self.open = Open::None;
        self.write_buf()?;
        Ok(())
    }

    /// Write an attr, make sure name and value contain only allowed chars.
    /// For an escaping version use `attr_esc`
    pub fn attr(&mut self, name: &str, value: &str) -> Result {
        if cfg!(feature = "check_xml") && self.open == Open::None {
            panic!("Attempted to write attr to elem, when no elem was opened, stack {:?}", self.stack);
        }
        self.buf.push(' ');
        self.buf.push_str(name);
        self.buf.push('=');
        self.buf.push('"');
        self.buf.push_str(value);
        self.buf.push('"');
        Ok(())
    }

    /// Write an attr, make sure name contains only allowed chars
    pub fn attr_esc(&mut self, name: &str, value: &str) -> Result {
        if cfg!(feature = "check_xml") && self.open == Open::None {
            panic!("Attempted to write attr to elem, when no elem was opened, stack {:?}", self.stack);
        }
        self.buf.push(' ');
        self.escape(name, true);
        self.buf.push('=');
        self.buf.push('"');
        self.escape(value, false);
        self.buf.push('"');
        Ok(())
    }

    /// Escape identifiers or text
    fn escape(&mut self, text: &str, ident: bool) {
        for c in text.chars() {
            match c {
                '"' => self.buf.push_str("&quot;"),
                '\'' => self.buf.push_str("&apos;"),
                '&' => self.buf.push_str("&amp;"),
                '<' => self.buf.push_str("&lt;"),
                '>' => self.buf.push_str("&gt;"),
                '\\' if ident => {
                    self.buf.push('\\');
                    self.buf.push('\\');
                }
                _ => {
                    self.buf.push(c);
                }
            };
        }
    }

    /// Write a text, doesn't escape the text.
    pub fn text(&mut self, text: &str) -> Result {
        self.close_elem()?;
        self.buf.push_str(text);
        Ok(())
    }

    /// Write a text, escapes the text automatically
    pub fn text_esc(&mut self, text: &str) -> Result {
        self.close_elem()?;
        self.escape(text, false);
        Ok(())
    }

    /// End and elem
    pub fn end_elem(&mut self, name: &str) -> Result {
        self.close_elem()?;

        if cfg!(feature = "check_xml") {
            match self.stack.pop() {
                Some(test) => {
                    if name != test {
                        panic!("Attempted to close elem {} but the open was {}, stack {:?}", name, test, self.stack)
                    }
                }
                None => panic!("Attempted to close an elem, when none was open, stack {:?}", self.stack)
            }
        }

        self.buf.push('<');
        self.buf.push('/');
        self.buf.push_str(name);
        self.buf.push('>');

        Ok(())
    }

    fn write_buf(&mut self) -> Result {
        self.writer.write_all(self.buf.as_bytes())?;
        self.buf.clear();
        Ok(())
    }

    /// Fails if there are any open elements.
    pub fn close(&mut self) -> Result {
        self.write_buf()?;

        if cfg!(feature = "check_xml") && !self.stack.is_empty() {
            panic!("Attempted to close the xml, but there are open elements on the stack {:?}", self.stack)
        }
        Ok(())
    }
}