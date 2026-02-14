use serde::de::IgnoredAny;
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
struct DocumentXml {
    #[serde(rename = "body", alias = "w:body", default)]
    body: DocumentBodyXml,
}

#[derive(Debug, Deserialize, Default)]
struct DocumentBodyXml {
    #[serde(rename = "$value", default)]
    children: Vec<DocumentChildXml>,
}

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
enum DocumentChildXml {
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
    StructuredDataTag(IgnoredAny),
    #[serde(rename = "sectPr", alias = "w:sectPr")]
    SectionProperty(SectionProperty),
    #[serde(other)]
    Unknown,
}

fn parse_optional_usize_doc(v: Option<String>) -> Option<usize> {
    v.and_then(|s| s.parse::<usize>().ok())
}

fn document_child_from_xml(xml: DocumentChildXml) -> Option<DocumentChild> {
    match xml {
        DocumentChildXml::Paragraph(p) => Some(DocumentChild::Paragraph(Box::new(p))),
        DocumentChildXml::Table(t) => Some(DocumentChild::Table(Box::new(t))),
        DocumentChildXml::BookmarkStart(node) => {
            let id = parse_optional_usize_doc(node.id)?;
            let name = node.name?;
            Some(DocumentChild::BookmarkStart(BookmarkStart::new(id, name)))
        }
        DocumentChildXml::BookmarkEnd(node) => {
            let id = parse_optional_usize_doc(node.id)?;
            Some(DocumentChild::BookmarkEnd(BookmarkEnd::new(id)))
        }
        DocumentChildXml::CommentStart(node) => {
            let id = parse_optional_usize_doc(node.id)?;
            Some(DocumentChild::CommentStart(Box::new(CommentRangeStart::new(
                Comment::new(id),
            ))))
        }
        DocumentChildXml::CommentEnd(node) => {
            let id = parse_optional_usize_doc(node.id)?;
            Some(DocumentChild::CommentEnd(CommentRangeEnd::new(id)))
        }
        DocumentChildXml::StructuredDataTag(_) | DocumentChildXml::Unknown => None,
        DocumentChildXml::SectionProperty(_) => None, // handled separately
    }
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Document {
    pub children: Vec<DocumentChild>,
    pub section_property: SectionProperty,
    pub has_numbering: bool,
}

impl<'de> Deserialize<'de> for Document {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = DocumentXml::deserialize(deserializer)?;
        let mut children = Vec::new();
        let mut has_numbering = false;
        let mut section_property = SectionProperty::new();

        for child in xml.body.children {
            match child {
                DocumentChildXml::SectionProperty(sp) => {
                    section_property = sp;
                }
                other => {
                    if let Some(mapped) = document_child_from_xml(other) {
                        match &mapped {
                            DocumentChild::Paragraph(p) => {
                                if p.has_numbering {
                                    has_numbering = true;
                                }
                            }
                            DocumentChild::Table(t) => {
                                if t.has_numbering {
                                    has_numbering = true;
                                }
                            }
                            _ => {}
                        }
                        children.push(mapped);
                    }
                }
            }
        }

        Ok(Document {
            children,
            section_property,
            has_numbering,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DocumentChild {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
    BookmarkStart(BookmarkStart),
    BookmarkEnd(BookmarkEnd),
    CommentStart(Box<CommentRangeStart>),
    CommentEnd(CommentRangeEnd),
    StructuredDataTag(Box<StructuredDataTag>),
    TableOfContents(Box<TableOfContents>),
    Section(Box<Section>),
}

impl Serialize for DocumentChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            DocumentChild::Paragraph(ref p) => {
                let mut t = serializer.serialize_struct("Paragraph", 2)?;
                t.serialize_field("type", "paragraph")?;
                t.serialize_field("data", p)?;
                t.end()
            }
            DocumentChild::Table(ref c) => {
                let mut t = serializer.serialize_struct("Table", 2)?;
                t.serialize_field("type", "table")?;
                t.serialize_field("data", c)?;
                t.end()
            }
            DocumentChild::BookmarkStart(ref c) => {
                let mut t = serializer.serialize_struct("BookmarkStart", 2)?;
                t.serialize_field("type", "bookmarkStart")?;
                t.serialize_field("data", c)?;
                t.end()
            }
            DocumentChild::BookmarkEnd(ref c) => {
                let mut t = serializer.serialize_struct("BookmarkEnd", 2)?;
                t.serialize_field("type", "bookmarkEnd")?;
                t.serialize_field("data", c)?;
                t.end()
            }
            DocumentChild::CommentStart(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeStart", 2)?;
                t.serialize_field("type", "commentRangeStart")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            DocumentChild::CommentEnd(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeEnd", 2)?;
                t.serialize_field("type", "commentRangeEnd")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            DocumentChild::StructuredDataTag(ref r) => {
                let mut t = serializer.serialize_struct("StructuredDataTag", 2)?;
                t.serialize_field("type", "structuredDataTag")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            DocumentChild::TableOfContents(ref r) => {
                let mut t = serializer.serialize_struct("TableOfContents", 2)?;
                t.serialize_field("type", "tableOfContents")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            DocumentChild::Section(ref r) => {
                let mut t = serializer.serialize_struct("Section", 2)?;
                t.serialize_field("type", "section")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

impl Default for Document {
    fn default() -> Self {
        Self {
            children: Vec::new(),
            section_property: SectionProperty::new(),
            has_numbering: false,
        }
    }
}

impl Document {
    pub fn new() -> Document {
        Default::default()
    }

    pub fn add_paragraph(mut self, p: Paragraph) -> Self {
        if p.has_numbering {
            self.has_numbering = true
        }
        self.children.push(DocumentChild::Paragraph(Box::new(p)));
        self
    }

    pub fn add_table(mut self, t: Table) -> Self {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children.push(DocumentChild::Table(Box::new(t)));
        self
    }

    pub fn add_bookmark_start(mut self, id: usize, name: impl Into<String>) -> Self {
        self.children
            .push(DocumentChild::BookmarkStart(BookmarkStart::new(id, name)));
        self
    }

    pub fn add_bookmark_end(mut self, id: usize) -> Self {
        self.children
            .push(DocumentChild::BookmarkEnd(BookmarkEnd::new(id)));
        self
    }

    pub fn add_comment_start(mut self, comment: Comment) -> Self {
        self.children.push(DocumentChild::CommentStart(Box::new(
            CommentRangeStart::new(comment),
        )));
        self
    }

    pub fn add_comment_end(mut self, id: usize) -> Self {
        self.children
            .push(DocumentChild::CommentEnd(CommentRangeEnd::new(id)));
        self
    }

    pub fn add_section(mut self, sec: Section) -> Self {
        self.children.push(DocumentChild::Section(Box::new(sec)));
        self
    }

    pub fn title_pg(mut self) -> Self {
        self.section_property = self.section_property.title_pg();
        self
    }

    pub fn page_size(mut self, size: PageSize) -> Self {
        self.section_property = self.section_property.page_size(size);
        self
    }

    pub fn page_margin(mut self, margin: crate::types::PageMargin) -> Self {
        self.section_property = self.section_property.page_margin(margin);
        self
    }

    pub fn page_orient(mut self, o: crate::types::PageOrientationType) -> Self {
        self.section_property = self.section_property.page_orient(o);
        self
    }

    pub fn doc_grid(mut self, doc_grid: DocGrid) -> Self {
        self.section_property = self.section_property.doc_grid(doc_grid);
        self
    }

    pub fn default_section_property(mut self, property: SectionProperty) -> Self {
        self.section_property = property;
        self
    }

    pub fn header(mut self, h: Header, rid: &str) -> Self {
        self.section_property = self.section_property.header(h, rid);
        self
    }

    pub fn first_header(mut self, h: Header, rid: &str) -> Self {
        self.section_property = self.section_property.first_header(h, rid);
        self
    }

    pub(crate) fn first_header_without_title_pg(mut self, h: Header, rid: &str) -> Self {
        self.section_property = self.section_property.first_header_without_title_pg(h, rid);
        self
    }

    pub fn even_header(mut self, h: Header, rid: &str) -> Self {
        self.section_property = self.section_property.even_header(h, rid);
        self
    }

    pub fn footer(mut self, h: Footer, rid: &str) -> Self {
        self.section_property = self.section_property.footer(h, rid);
        self
    }

    pub fn first_footer(mut self, h: Footer, rid: &str) -> Self {
        self.section_property = self.section_property.first_footer(h, rid);
        self
    }

    pub(crate) fn first_footer_without_title_pg(mut self, h: Footer, rid: &str) -> Self {
        self.section_property = self.section_property.first_footer_without_title_pg(h, rid);
        self
    }

    pub fn even_footer(mut self, h: Footer, rid: &str) -> Self {
        self.section_property = self.section_property.even_footer(h, rid);
        self
    }

    pub fn add_structured_data_tag(mut self, t: StructuredDataTag) -> Self {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children
            .push(DocumentChild::StructuredDataTag(Box::new(t)));
        self
    }

    pub fn add_table_of_contents(mut self, t: TableOfContents) -> Self {
        self.children
            .push(DocumentChild::TableOfContents(Box::new(t)));
        self
    }

    pub fn columns(mut self, col: usize) -> Self {
        self.section_property.columns = col;
        self
    }

    pub fn text_direction(mut self, direction: String) -> Self {
        self.section_property.text_direction = direction;
        self
    }

    pub fn page_num_type(mut self, p: PageNumType) -> Self {
        self.section_property = self.section_property.page_num_type(p);
        self
    }
}

impl BuildXML for DocumentChild {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        match self {
            DocumentChild::Paragraph(v) => v.build_to(stream),
            DocumentChild::Table(v) => v.build_to(stream),
            DocumentChild::BookmarkStart(v) => v.build_to(stream),
            DocumentChild::BookmarkEnd(v) => v.build_to(stream),
            DocumentChild::CommentStart(v) => v.build_to(stream),
            DocumentChild::CommentEnd(v) => v.build_to(stream),
            DocumentChild::StructuredDataTag(v) => v.build_to(stream),
            DocumentChild::TableOfContents(v) => v.build_to(stream),
            DocumentChild::Section(v) => v.build_to(stream),
        }
    }
}

impl BuildXML for Document {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .declaration(Some(true))?
            .open_document()?
            .open_body()?
            .add_children(&self.children)?
            .add_child(&self.section_property)?
            .close()?
            .close()?
            .into_inner()
    }
}

#[cfg(test)]
mod tests {

    use super::super::Run;
    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_document() {
        let b = Document::new()
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Hello")))
            .build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 wp14"><w:body><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr><w:r><w:rPr /><w:t xml:space="preserve">Hello</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="1" /></w:sectPr></w:body></w:document>"#
        );
    }

    #[test]
    fn test_document_with_toc() {
        let toc = TableOfContents::new().heading_styles_range(1, 3);
        let b = Document::new().add_table_of_contents(toc).build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 wp14"><w:body><w:sdt><w:sdtPr><w:rPr /></w:sdtPr><w:sdtContent><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr><w:r><w:rPr /><w:fldChar w:fldCharType="begin" w:dirty="true" /><w:instrText>TOC \o &quot;1-3&quot;</w:instrText><w:fldChar w:fldCharType="separate" w:dirty="false" /></w:r></w:p><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr><w:r><w:rPr /><w:fldChar w:fldCharType="end" w:dirty="false" /></w:r></w:p></w:sdtContent></w:sdt><w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="1" /></w:sectPr></w:body></w:document>"#
        );
    }

    #[test]
    fn test_document_cols() {
        let b = Document::new()
            .columns(2)
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Hello")))
            .build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 wp14"><w:body><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr><w:r><w:rPr /><w:t xml:space="preserve">Hello</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="2" /></w:sectPr></w:body></w:document>"#
        );
    }

    // XML Deserialization tests (quick-xml serde)
    #[test]
    fn test_document_xml_deserialize_children() {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:pPr />
                    <w:r><w:t>Hello</w:t></w:r>
                </w:p>
                <w:bookmarkStart w:id="1" w:name="bm" />
                <w:commentRangeStart w:id="2" />
                <w:commentRangeEnd w:id="2" />
                <w:bookmarkEnd w:id="1" />
            </w:body>
        </w:document>"#;

        let doc: Document = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(doc.children.len(), 5);
        assert!(matches!(&doc.children[0], DocumentChild::Paragraph(_)));
        assert!(matches!(&doc.children[1], DocumentChild::BookmarkStart(_)));
        assert!(matches!(&doc.children[2], DocumentChild::CommentStart(_)));
        assert!(matches!(&doc.children[3], DocumentChild::CommentEnd(_)));
        assert!(matches!(&doc.children[4], DocumentChild::BookmarkEnd(_)));
        assert!(!doc.has_numbering);
    }

    #[test]
    fn test_document_xml_deserialize_has_numbering() {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:pPr>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
                    </w:pPr>
                </w:p>
            </w:body>
        </w:document>"#;

        let doc: Document = quick_xml::de::from_str(xml).unwrap();
        assert!(doc.has_numbering);
    }
}
