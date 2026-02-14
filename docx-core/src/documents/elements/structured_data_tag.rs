use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use super::*;
use crate::documents::BuildXML;
// use crate::types::*;
use crate::xml_builder::*;

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct XmlBookmarkStartNode {
    #[serde(rename = "@id", alias = "@w:id", default)]
    id: Option<String>,
    #[serde(rename = "@name", alias = "@w:name", default)]
    name: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlIdNode {
    #[serde(rename = "@id", alias = "@w:id", default)]
    id: Option<String>,
}

#[derive(Debug, Deserialize)]
enum SdtContentChildXml {
    #[serde(rename = "r", alias = "w:r")]
    Run(Run),
    #[serde(rename = "p", alias = "w:p")]
    Paragraph(Paragraph),
    #[serde(rename = "tbl", alias = "w:tbl")]
    Table(Table),
    #[serde(rename = "bookmarkStart", alias = "w:bookmarkStart")]
    BookmarkStart(XmlBookmarkStartNode),
    #[serde(rename = "bookmarkEnd", alias = "w:bookmarkEnd")]
    BookmarkEnd(XmlIdNode),
    #[serde(rename = "commentRangeStart", alias = "w:commentRangeStart")]
    CommentStart(XmlIdNode),
    #[serde(rename = "commentRangeEnd", alias = "w:commentRangeEnd")]
    CommentEnd(XmlIdNode),
    #[serde(rename = "sdt", alias = "w:sdt")]
    StructuredDataTag(Box<StructuredDataTag>),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct SdtContentXml {
    #[serde(rename = "$value", default)]
    children: Vec<SdtContentChildXml>,
}

#[derive(Debug, Deserialize, Default)]
struct StructuredDataTagXml {
    #[serde(rename = "sdtPr", alias = "w:sdtPr", default)]
    property: Option<StructuredDataTagProperty>,
    #[serde(rename = "sdtContent", alias = "w:sdtContent", default)]
    content: Option<SdtContentXml>,
}

fn parse_optional_usize(v: Option<String>) -> Option<usize> {
    v.and_then(|s| s.parse::<usize>().ok())
}

fn sdt_child_from_xml(xml: SdtContentChildXml) -> Option<StructuredDataTagChild> {
    match xml {
        SdtContentChildXml::Run(run) => Some(StructuredDataTagChild::Run(Box::new(run))),
        SdtContentChildXml::Paragraph(p) => {
            Some(StructuredDataTagChild::Paragraph(Box::new(p)))
        }
        SdtContentChildXml::Table(t) => Some(StructuredDataTagChild::Table(Box::new(t))),
        SdtContentChildXml::BookmarkStart(node) => {
            let id = parse_optional_usize(node.id)?;
            let name = node.name?;
            Some(StructuredDataTagChild::BookmarkStart(BookmarkStart::new(
                id, name,
            )))
        }
        SdtContentChildXml::BookmarkEnd(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(StructuredDataTagChild::BookmarkEnd(BookmarkEnd::new(id)))
        }
        SdtContentChildXml::CommentStart(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(StructuredDataTagChild::CommentStart(Box::new(
                CommentRangeStart::new(Comment::new(id)),
            )))
        }
        SdtContentChildXml::CommentEnd(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(StructuredDataTagChild::CommentEnd(CommentRangeEnd::new(id)))
        }
        SdtContentChildXml::StructuredDataTag(sdt) => {
            Some(StructuredDataTagChild::StructuredDataTag(sdt))
        }
        SdtContentChildXml::Unknown => None,
    }
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct StructuredDataTag {
    pub children: Vec<StructuredDataTagChild>,
    pub property: StructuredDataTagProperty,
    pub has_numbering: bool,
}

impl<'de> Deserialize<'de> for StructuredDataTag {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = StructuredDataTagXml::deserialize(deserializer)?;
        let children: Vec<StructuredDataTagChild> = xml
            .content
            .map(|c| c.children.into_iter().filter_map(sdt_child_from_xml).collect())
            .unwrap_or_default();

        let has_numbering = children.iter().any(|c| match c {
            StructuredDataTagChild::Paragraph(p) => p.has_numbering,
            StructuredDataTagChild::Table(t) => t.has_numbering,
            StructuredDataTagChild::StructuredDataTag(s) => s.has_numbering,
            _ => false,
        });

        Ok(StructuredDataTag {
            children,
            property: xml.property.unwrap_or_default(),
            has_numbering,
        })
    }
}

impl Default for StructuredDataTag {
    fn default() -> Self {
        Self {
            children: Vec::new(),
            property: StructuredDataTagProperty::new(),
            has_numbering: false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum StructuredDataTagChild {
    Run(Box<Run>),
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
    BookmarkStart(BookmarkStart),
    BookmarkEnd(BookmarkEnd),
    CommentStart(Box<CommentRangeStart>),
    CommentEnd(CommentRangeEnd),
    StructuredDataTag(Box<StructuredDataTag>),
}

impl BuildXML for StructuredDataTagChild {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        match self {
            StructuredDataTagChild::Run(v) => v.build_to(stream),
            StructuredDataTagChild::Paragraph(v) => v.build_to(stream),
            StructuredDataTagChild::Table(v) => v.build_to(stream),
            StructuredDataTagChild::BookmarkStart(v) => v.build_to(stream),
            StructuredDataTagChild::BookmarkEnd(v) => v.build_to(stream),
            StructuredDataTagChild::CommentStart(v) => v.build_to(stream),
            StructuredDataTagChild::CommentEnd(v) => v.build_to(stream),
            StructuredDataTagChild::StructuredDataTag(v) => v.build_to(stream),
        }
    }
}

impl Serialize for StructuredDataTagChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            StructuredDataTagChild::Run(ref r) => {
                let mut t = serializer.serialize_struct("Run", 2)?;
                t.serialize_field("type", "run")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            StructuredDataTagChild::Paragraph(ref r) => {
                let mut t = serializer.serialize_struct("Paragraph", 2)?;
                t.serialize_field("type", "paragraph")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            StructuredDataTagChild::Table(ref r) => {
                let mut t = serializer.serialize_struct("Table", 2)?;
                t.serialize_field("type", "table")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            StructuredDataTagChild::BookmarkStart(ref c) => {
                let mut t = serializer.serialize_struct("BookmarkStart", 2)?;
                t.serialize_field("type", "bookmarkStart")?;
                t.serialize_field("data", c)?;
                t.end()
            }
            StructuredDataTagChild::BookmarkEnd(ref c) => {
                let mut t = serializer.serialize_struct("BookmarkEnd", 2)?;
                t.serialize_field("type", "bookmarkEnd")?;
                t.serialize_field("data", c)?;
                t.end()
            }
            StructuredDataTagChild::CommentStart(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeStart", 2)?;
                t.serialize_field("type", "commentRangeStart")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            StructuredDataTagChild::CommentEnd(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeEnd", 2)?;
                t.serialize_field("type", "commentRangeEnd")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            StructuredDataTagChild::StructuredDataTag(ref r) => {
                let mut t = serializer.serialize_struct("StructuredDataTag", 2)?;
                t.serialize_field("type", "structuredDataTag")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

impl StructuredDataTag {
    pub fn new() -> Self {
        Default::default()
    }

    pub fn add_run(mut self, run: Run) -> Self {
        self.children
            .push(StructuredDataTagChild::Run(Box::new(run)));
        self
    }

    pub fn add_paragraph(mut self, p: Paragraph) -> Self {
        if p.has_numbering {
            self.has_numbering = true
        }
        self.children
            .push(StructuredDataTagChild::Paragraph(Box::new(p)));
        self
    }

    pub fn add_table(mut self, t: Table) -> Self {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children
            .push(StructuredDataTagChild::Table(Box::new(t)));
        self
    }

    pub fn data_binding(mut self, d: DataBinding) -> Self {
        self.property = self.property.data_binding(d);
        self
    }

    pub fn alias(mut self, v: impl Into<String>) -> Self {
        self.property = self.property.alias(v);
        self
    }
}

impl BuildXML for StructuredDataTag {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_structured_tag()?
            .add_child(&self.property)?
            .open_structured_tag_content()?
            .add_children(&self.children)?
            .close()?
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
    fn test_sdt() {
        let b = StructuredDataTag::new()
            .data_binding(DataBinding::new().xpath("root/hello"))
            .add_run(Run::new().add_text("Hello"))
            .build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:sdt><w:sdtPr><w:rPr /><w:dataBinding w:xpath="root/hello" /></w:sdtPr><w:sdtContent><w:r><w:rPr /><w:t xml:space="preserve">Hello</w:t></w:r></w:sdtContent></w:sdt>"#
        );
    }

    #[test]
    fn test_sdt_xml_deserialize() {
        let xml = r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:sdtPr>
                <w:rPr />
                <w:alias w:val="Test SDT" />
                <w:dataBinding w:xpath="root/data" />
            </w:sdtPr>
            <w:sdtContent>
                <w:p />
                <w:r><w:t>Text content</w:t></w:r>
            </w:sdtContent>
        </w:sdt>"#;

        let sdt: StructuredDataTag = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(sdt.property.alias, Some("Test SDT".to_string()));
        assert!(sdt.property.data_binding.is_some());
        assert_eq!(sdt.children.len(), 2);
        assert!(matches!(&sdt.children[0], StructuredDataTagChild::Paragraph(_)));
        assert!(matches!(&sdt.children[1], StructuredDataTagChild::Run(_)));
    }

    #[test]
    fn test_sdt_xml_deserialize_nested() {
        let xml = r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:sdtPr />
            <w:sdtContent>
                <w:sdt>
                    <w:sdtPr><w:alias w:val="Nested" /></w:sdtPr>
                    <w:sdtContent>
                        <w:p />
                    </w:sdtContent>
                </w:sdt>
            </w:sdtContent>
        </w:sdt>"#;

        let sdt: StructuredDataTag = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(sdt.children.len(), 1);
        if let StructuredDataTagChild::StructuredDataTag(nested) = &sdt.children[0] {
            assert_eq!(nested.property.alias, Some("Nested".to_string()));
            assert_eq!(nested.children.len(), 1);
        } else {
            panic!("Expected nested StructuredDataTag");
        }
    }
}
