use super::*;
use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use crate::documents::BuildXML;
use crate::xml_builder::*;

#[derive(Debug, Clone, Serialize, PartialEq, Default)]
pub struct TextBoxContent {
    pub children: Vec<TextBoxContentChild>,
    pub has_numbering: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextBoxContentChild {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
}

// ============================================================================
// XML Deserialization (quick-xml serde)
// ============================================================================

#[derive(Deserialize)]
enum TextBoxContentChildXml {
    #[serde(rename = "p", alias = "w:p")]
    Paragraph(Paragraph),
    #[serde(rename = "tbl", alias = "w:tbl")]
    Table(Table),
    #[serde(other)]
    Unknown,
}

#[derive(Deserialize)]
struct TextBoxContentXml {
    #[serde(rename = "$value", default)]
    children: Vec<TextBoxContentChildXml>,
}

impl<'de> Deserialize<'de> for TextBoxContent {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = TextBoxContentXml::deserialize(deserializer)?;
        let mut has_numbering = false;
        let children = xml
            .children
            .into_iter()
            .filter_map(|c| match c {
                TextBoxContentChildXml::Paragraph(p) => {
                    if p.has_numbering {
                        has_numbering = true;
                    }
                    Some(TextBoxContentChild::Paragraph(Box::new(p)))
                }
                TextBoxContentChildXml::Table(t) => {
                    if t.has_numbering {
                        has_numbering = true;
                    }
                    Some(TextBoxContentChild::Table(Box::new(t)))
                }
                TextBoxContentChildXml::Unknown => None,
            })
            .collect();
        Ok(TextBoxContent {
            children,
            has_numbering,
        })
    }
}

impl Serialize for TextBoxContentChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            TextBoxContentChild::Paragraph(ref p) => {
                let mut t = serializer.serialize_struct("Paragraph", 2)?;
                t.serialize_field("type", "paragraph")?;
                t.serialize_field("data", p)?;
                t.end()
            }
            TextBoxContentChild::Table(ref c) => {
                let mut t = serializer.serialize_struct("Table", 2)?;
                t.serialize_field("type", "table")?;
                t.serialize_field("data", c)?;
                t.end()
            }
        }
    }
}

impl TextBoxContent {
    pub fn new() -> TextBoxContent {
        Default::default()
    }

    pub fn add_paragraph(mut self, p: Paragraph) -> Self {
        if p.has_numbering {
            self.has_numbering = true
        }
        self.children
            .push(TextBoxContentChild::Paragraph(Box::new(p)));
        self
    }

    pub fn add_table(mut self, t: Table) -> Self {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children.push(TextBoxContentChild::Table(Box::new(t)));
        self
    }
}

impl BuildXML for TextBoxContentChild {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        match self {
            TextBoxContentChild::Paragraph(p) => p.build_to(stream),
            TextBoxContentChild::Table(t) => t.build_to(stream),
        }
    }
}

impl BuildXML for TextBoxContent {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_text_box_content()?
            .add_children(&self.children)?
            .close()?
            .into_inner()
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_text_box_content_build() {
        let b = TextBoxContent::new()
            .add_paragraph(Paragraph::new())
            .build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:txbxContent><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr></w:p></w:txbxContent>"#
        );
    }
}
