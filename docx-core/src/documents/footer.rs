use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use super::*;
use crate::documents::BuildXML;
use crate::xml_builder::*;

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct FooterXml {
    #[serde(rename = "$value", default)]
    children: Vec<FooterChildXml>,
}

#[derive(Debug, Deserialize)]
enum FooterChildXml {
    #[serde(rename = "p", alias = "w:p")]
    Paragraph(Paragraph),
    #[serde(rename = "tbl", alias = "w:tbl")]
    Table(Table),
    #[serde(rename = "sdt", alias = "w:sdt")]
    StructuredDataTag(StructuredDataTag),
    #[serde(other)]
    Unknown,
}

fn footer_child_from_xml(xml: FooterChildXml) -> Option<FooterChild> {
    match xml {
        FooterChildXml::Paragraph(p) => Some(FooterChild::Paragraph(Box::new(p))),
        FooterChildXml::Table(t) => Some(FooterChild::Table(Box::new(t))),
        FooterChildXml::StructuredDataTag(sdt) => {
            Some(FooterChild::StructuredDataTag(Box::new(sdt)))
        }
        FooterChildXml::Unknown => None,
    }
}

impl<'de> Deserialize<'de> for Footer {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = FooterXml::deserialize(deserializer)?;
        let mut children = Vec::new();
        let mut has_numbering = false;

        for child in xml.children {
            if let Some(child) = footer_child_from_xml(child) {
                if matches!(&child, FooterChild::Paragraph(p) if p.has_numbering)
                    || matches!(&child, FooterChild::Table(t) if t.has_numbering)
                {
                    has_numbering = true;
                }
                children.push(child);
            }
        }

        Ok(Footer {
            has_numbering,
            children,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Footer {
    pub has_numbering: bool,
    pub children: Vec<FooterChild>,
}

impl Footer {
    pub fn new() -> Footer {
        Default::default()
    }

    pub fn add_paragraph(mut self, p: Paragraph) -> Self {
        if p.has_numbering {
            self.has_numbering = true
        }
        self.children.push(FooterChild::Paragraph(Box::new(p)));
        self
    }

    pub fn add_table(mut self, t: Table) -> Self {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children.push(FooterChild::Table(Box::new(t)));
        self
    }

    /// reader only
    pub(crate) fn add_structured_data_tag(mut self, t: StructuredDataTag) -> Self {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children
            .push(FooterChild::StructuredDataTag(Box::new(t)));
        self
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FooterChild {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
    StructuredDataTag(Box<StructuredDataTag>),
}

impl Serialize for FooterChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            FooterChild::Paragraph(ref p) => {
                let mut t = serializer.serialize_struct("Paragraph", 2)?;
                t.serialize_field("type", "paragraph")?;
                t.serialize_field("data", p)?;
                t.end()
            }
            FooterChild::Table(ref c) => {
                let mut t = serializer.serialize_struct("Table", 2)?;
                t.serialize_field("type", "table")?;
                t.serialize_field("data", c)?;
                t.end()
            }
            FooterChild::StructuredDataTag(ref r) => {
                let mut t = serializer.serialize_struct("StructuredDataTag", 2)?;
                t.serialize_field("type", "structuredDataTag")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

impl BuildXML for Footer {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .declaration(Some(true))?
            .open_footer()?
            .apply_each(&self.children, |c, b| match c {
                FooterChild::Paragraph(p) => b.add_child(&p),
                FooterChild::Table(t) => b.add_child(&t),
                FooterChild::StructuredDataTag(t) => b.add_child(&t),
            })?
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
    fn test_settings() {
        let c = Footer::new();
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14" />"#
        );
    }
}
