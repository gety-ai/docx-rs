// Licensed under either of
//
// Apache License, Version 2.0, (LICENSE-APACHE or http://www.apache.org/licenses/LICENSE-2.0)
// MIT license (LICENSE-MIT or http://opensource.org/licenses/MIT)
// at your option.
//
// Contribution
// Unless you explicitly state otherwise, any contribution intentionally submitted for inclusion in the work by you, as defined in the Apache-2.0 license, shall be dual licensed as above, without any additional terms or conditions.
use quick_xml::events::Event;
use quick_xml::Reader;
use serde::Serialize;
use std::fmt::{Display, Formatter, Write};
use std::io::{BufReader, Cursor, Read};
use std::str::FromStr;

/// An XML Document
#[derive(Debug, Clone)]
pub struct XmlDocument {
    /// Data contained within the parsed XML Document
    pub data: Vec<XmlData>,
}

impl Display for XmlDocument {
    fn fmt(&self, f: &mut Formatter) -> std::fmt::Result {
        for item in self.data.iter() {
            item.fmt(f)?;
        }
        Ok(())
    }
}

/// An XML Tag
///
/// For example:
///
/// ```XML
/// <foo bar="baz">
///     test text
///     <sub></sub>
/// </foo>
/// ```
#[derive(Debug, Clone, Serialize)]
pub struct XmlData {
    /// Name of the tag (i.e. "foo")
    pub name: String,
    /// Key-value pairs of the attributes (i.e. ("bar", "baz"))
    pub attributes: Vec<(String, String)>,
    /// Data (i.e. "test text")
    pub data: Option<String>,
    /// Sub elements (i.e. an XML element of "sub")
    pub children: Vec<XmlData>,
}

impl XmlData {
    /// Format the XML data as a string
    fn format(self: &XmlData, f: &mut Formatter, _depth: usize) -> std::fmt::Result {
        write!(f, "<{}", self.name)?;

        for (key, val) in self.attributes.iter() {
            write!(f, r#" {}="{}""#, key, val)?;
        }

        f.write_char('>')?;

        if let Some(ref data) = self.data {
            write!(f, "{}", data)?
        }

        for child in self.children.iter() {
            child.format(f, _depth + 1)?;
        }

        write!(f, "</{}>", self.name)
    }
}

impl Display for XmlData {
    fn fmt(&self, f: &mut Formatter) -> std::fmt::Result {
        self.format(f, 0)
    }
}

fn read_element(
    e: &quick_xml::events::BytesStart,
) -> Result<(String, Vec<(String, String)>), ParseXmlError> {
    let name = std::str::from_utf8(e.name().as_ref())
        .map_err(|e| ParseXmlError(e.to_string()))?
        .to_string();
    let attributes = e
        .attributes()
        .map(|a| {
            let a = a.map_err(|e| ParseXmlError(format!("{e}")))?;
            let key = std::str::from_utf8(a.key.as_ref())
                .map_err(|e| ParseXmlError(e.to_string()))?
                .to_string();
            let val_bytes: &[u8] = &a.value;
            let val_str = std::str::from_utf8(val_bytes)
                .map_err(|e| ParseXmlError(e.to_string()))?;
            let val = quick_xml::escape::unescape(val_str)
                .map_err(|e| ParseXmlError(e.to_string()))?
                .to_string();
            Ok((key, val))
        })
        .collect::<Result<Vec<_>, ParseXmlError>>()?;
    Ok((name, attributes))
}

impl XmlDocument {
    pub fn from_reader<R>(source: R, trim: bool) -> Result<Self, ParseXmlError>
    where
        R: Read,
    {
        let mut reader = Reader::from_reader(BufReader::new(source));
        let mut buf = Vec::new();
        let mut stack: Vec<XmlData> = Vec::new();
        let mut root_items: Vec<XmlData> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let (name, attributes) = read_element(e)?;
                    stack.push(XmlData {
                        name,
                        attributes,
                        data: None,
                        children: Vec::new(),
                    });
                }
                Ok(Event::End(ref e)) => {
                    let name_bytes = e.name();
                    let end_name = std::str::from_utf8(name_bytes.as_ref())
                        .map_err(|e| ParseXmlError(e.to_string()))?;
                    let node = stack
                        .pop()
                        .ok_or_else(|| ParseXmlError(format!("Invalid end tag: {end_name}")))?;
                    if node.name != end_name {
                        return Err(ParseXmlError(format!(
                            "Invalid end tag: expected {}, got {end_name}",
                            node.name
                        )));
                    }
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    } else {
                        root_items.push(node);
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    let (name, attributes) = read_element(e)?;
                    let node = XmlData {
                        name,
                        attributes,
                        data: None,
                        children: Vec::new(),
                    };
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    } else {
                        root_items.push(node);
                    }
                }
                Ok(Event::Text(ref t)) => {
                    let text = t
                        .unescape()
                        .map_err(|e| ParseXmlError(e.to_string()))?;
                    let text = if trim {
                        text.trim().to_string()
                    } else {
                        text.to_string()
                    };
                    if let Some(current) = stack.last_mut() {
                        current.data = Some(text);
                    }
                }
                Ok(Event::CData(ref t)) => {
                    let text = String::from_utf8(t.to_vec())
                        .map_err(|e| ParseXmlError(e.to_string()))?;
                    let text = if trim {
                        text.trim().to_string()
                    } else {
                        text
                    };
                    if let Some(current) = stack.last_mut() {
                        current.data = Some(text);
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => return Err(ParseXmlError(e.to_string())),
            }
            buf.clear();
        }

        if !stack.is_empty() {
            return Err(ParseXmlError("Invalid end tag".to_string()));
        }

        Ok(XmlDocument { data: root_items })
    }
}

/// Error when parsing XML
#[derive(Debug, Clone, PartialEq)]
pub struct ParseXmlError(String);

impl Display for ParseXmlError {
    fn fmt(&self, f: &mut Formatter) -> std::fmt::Result {
        write!(f, "Coult not parse string to XML: {}", self.0)
    }
}

// Generate an XML document from a string
impl FromStr for XmlDocument {
    type Err = ParseXmlError;

    fn from_str(s: &str) -> Result<XmlDocument, ParseXmlError> {
        XmlDocument::from_reader(Cursor::new(s.to_string().into_bytes()), true)
    }
}
