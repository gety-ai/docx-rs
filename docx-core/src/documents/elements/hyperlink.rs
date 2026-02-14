use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use super::*;
use crate::documents::BuildXML;
use crate::escape::escape;
use crate::types::*;
use crate::{create_hyperlink_rid, generate_hyperlink_id, xml_builder::*};

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
enum HyperlinkChildXml {
    #[serde(rename = "r", alias = "w:r")]
    Run(Run),
    #[serde(rename = "bookmarkStart", alias = "w:bookmarkStart")]
    BookmarkStart(XmlBookmarkStartNode),
    #[serde(rename = "bookmarkEnd", alias = "w:bookmarkEnd")]
    BookmarkEnd(XmlIdNode),
    #[serde(rename = "ins", alias = "w:ins")]
    Insert(Insert),
    #[serde(rename = "del", alias = "w:del")]
    Delete(Delete),
    #[serde(rename = "commentRangeStart", alias = "w:commentRangeStart")]
    CommentStart(XmlIdNode),
    #[serde(rename = "commentRangeEnd", alias = "w:commentRangeEnd")]
    CommentEnd(XmlIdNode),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct HyperlinkXml {
    #[serde(rename = "@id", alias = "@r:id", default)]
    rid: Option<String>,
    #[serde(rename = "@anchor", alias = "@w:anchor", default)]
    anchor: Option<String>,
    #[serde(rename = "@history", alias = "@w:history", default)]
    history: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<HyperlinkChildXml>,
}

fn parse_history(v: Option<String>) -> Option<usize> {
    v.and_then(|s| {
        let s = s.trim().to_lowercase();
        match s.as_str() {
            "on" | "true" | "1" => Some(1),
            "off" | "false" | "0" => Some(0),
            _ => s.parse::<usize>().ok(),
        }
    })
}

fn parse_optional_usize(v: Option<String>) -> Option<usize> {
    v.and_then(|s| s.parse::<usize>().ok())
}

fn hyperlink_child_from_xml(xml: HyperlinkChildXml) -> Option<ParagraphChild> {
    match xml {
        HyperlinkChildXml::Run(run) => Some(ParagraphChild::Run(Box::new(run))),
        HyperlinkChildXml::BookmarkStart(node) => {
            let id = parse_optional_usize(node.id)?;
            let name = node.name?;
            Some(ParagraphChild::BookmarkStart(BookmarkStart::new(id, name)))
        }
        HyperlinkChildXml::BookmarkEnd(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(ParagraphChild::BookmarkEnd(BookmarkEnd::new(id)))
        }
        HyperlinkChildXml::Insert(ins) => Some(ParagraphChild::Insert(ins)),
        HyperlinkChildXml::Delete(del) => Some(ParagraphChild::Delete(del)),
        HyperlinkChildXml::CommentStart(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(ParagraphChild::CommentStart(Box::new(
                CommentRangeStart::new(Comment::new(id)),
            )))
        }
        HyperlinkChildXml::CommentEnd(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(ParagraphChild::CommentEnd(CommentRangeEnd::new(id)))
        }
        HyperlinkChildXml::Unknown => None,
    }
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(tag = "type")]
#[serde(rename_all = "camelCase")]
pub enum HyperlinkData {
    External {
        rid: String,
        // path is writer only
        #[serde(skip_serializing_if = "String::is_empty")]
        path: String,
    },
    Anchor {
        anchor: String,
    },
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    #[serde(flatten)]
    pub link: HyperlinkData,
    pub history: Option<usize>,
    pub children: Vec<ParagraphChild>,
}

impl<'de> Deserialize<'de> for Hyperlink {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = HyperlinkXml::deserialize(deserializer)?;
        let link = if let Some(rid) = xml.rid.filter(|s| !s.is_empty()) {
            HyperlinkData::External {
                rid,
                path: String::default(),
            }
        } else if let Some(anchor) = xml.anchor.filter(|s| !s.is_empty()) {
            HyperlinkData::Anchor { anchor }
        } else {
            HyperlinkData::External {
                rid: String::default(),
                path: String::default(),
            }
        };

        Ok(Hyperlink {
            link,
            history: parse_history(xml.history),
            children: xml
                .children
                .into_iter()
                .filter_map(hyperlink_child_from_xml)
                .collect(),
        })
    }
}

impl Hyperlink {
    pub fn new(value: impl Into<String>, t: HyperlinkType) -> Self {
        let link = {
            match t {
                HyperlinkType::External => HyperlinkData::External {
                    rid: create_hyperlink_rid(generate_hyperlink_id()),
                    path: escape(&value.into()),
                },
                HyperlinkType::Anchor => HyperlinkData::Anchor {
                    anchor: value.into(),
                },
            }
        };
        Hyperlink {
            link,
            history: None,
            children: vec![],
        }
    }

    pub fn add_run(mut self, run: Run) -> Self {
        self.children.push(ParagraphChild::Run(Box::new(run)));
        self
    }

    pub fn add_structured_data_tag(mut self, t: StructuredDataTag) -> Self {
        self.children
            .push(ParagraphChild::StructuredDataTag(Box::new(t)));
        self
    }

    pub fn add_insert(mut self, insert: Insert) -> Self {
        self.children.push(ParagraphChild::Insert(insert));
        self
    }

    pub fn add_delete(mut self, delete: Delete) -> Self {
        self.children.push(ParagraphChild::Delete(delete));
        self
    }

    pub fn add_bookmark_start(mut self, id: usize, name: impl Into<String>) -> Self {
        self.children
            .push(ParagraphChild::BookmarkStart(BookmarkStart::new(id, name)));
        self
    }

    pub fn add_bookmark_end(mut self, id: usize) -> Self {
        self.children
            .push(ParagraphChild::BookmarkEnd(BookmarkEnd::new(id)));
        self
    }

    pub fn add_comment_start(mut self, comment: Comment) -> Self {
        self.children.push(ParagraphChild::CommentStart(Box::new(
            CommentRangeStart::new(comment),
        )));
        self
    }

    pub fn add_comment_end(mut self, id: usize) -> Self {
        self.children
            .push(ParagraphChild::CommentEnd(CommentRangeEnd::new(id)));
        self
    }
}

impl BuildXML for Hyperlink {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .apply(|b| match self.link {
                HyperlinkData::Anchor { ref anchor } => b.open_hyperlink(
                    None,
                    Some(anchor.clone()).as_ref(),
                    Some(self.history.unwrap_or(1)),
                ),
                HyperlinkData::External { ref rid, .. } => b.open_hyperlink(
                    Some(rid.clone()).as_ref(),
                    None,
                    Some(self.history.unwrap_or(1)),
                ),
            })?
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
    fn test_hyperlink() {
        let l =
            Hyperlink::new("ToC1", HyperlinkType::Anchor).add_run(Run::new().add_text("hello"));
        let b = l.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:hyperlink w:anchor="ToC1" w:history="1"><w:r><w:rPr /><w:t xml:space="preserve">hello</w:t></w:r></w:hyperlink>"#
        );
    }

    #[test]
    fn test_hyperlink_xml_deserialize_external() {
        let xml = r#"<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5" w:history="1">
            <w:r><w:t>click here</w:t></w:r>
            <w:bookmarkStart w:id="0" w:name="bm"/>
            <w:bookmarkEnd w:id="0"/>
        </w:hyperlink>"#;

        let link: Hyperlink = quick_xml::de::from_str(xml).unwrap();
        assert!(matches!(
            link.link,
            HyperlinkData::External { ref rid, ref path } if rid == "rId5" && path.is_empty()
        ));
        assert_eq!(link.history, Some(1));
        assert_eq!(link.children.len(), 3);
        assert!(matches!(&link.children[0], ParagraphChild::Run(_)));
        assert!(matches!(
            &link.children[1],
            ParagraphChild::BookmarkStart(b) if b == &BookmarkStart::new(0, "bm")
        ));
        assert!(matches!(
            &link.children[2],
            ParagraphChild::BookmarkEnd(b) if b == &BookmarkEnd::new(0)
        ));
    }

    #[test]
    fn test_hyperlink_xml_deserialize_anchor() {
        let xml = r#"<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:anchor="section1" w:history="1">
            <w:r><w:t>go to section</w:t></w:r>
        </w:hyperlink>"#;

        let link: Hyperlink = quick_xml::de::from_str(xml).unwrap();
        assert!(matches!(
            link.link,
            HyperlinkData::Anchor { ref anchor } if anchor == "section1"
        ));
        assert_eq!(link.history, Some(1));
        assert_eq!(link.children.len(), 1);
    }

    #[test]
    fn test_hyperlink_xml_deserialize_history_on_off() {
        // Test "on" value
        let xml_on = r#"<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:anchor="sec1" w:history="on">
            <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>"#;
        let link_on: Hyperlink = quick_xml::de::from_str(xml_on).unwrap();
        assert_eq!(link_on.history, Some(1));

        // Test "off" value
        let xml_off = r#"<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:anchor="sec2" w:history="off">
            <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>"#;
        let link_off: Hyperlink = quick_xml::de::from_str(xml_off).unwrap();
        assert_eq!(link_off.history, Some(0));

        // Test "true"/"false"
        let xml_true = r#"<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:anchor="sec3" w:history="true">
            <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>"#;
        let link_true: Hyperlink = quick_xml::de::from_str(xml_true).unwrap();
        assert_eq!(link_true.history, Some(1));
    }

    #[test]
    fn test_hyperlink_xml_deserialize_empty_rid() {
        // Test empty rid falls back to external with empty rid
        let xml = r#"<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="">
            <w:r><w:t>link</w:t></w:r>
        </w:hyperlink>"#;
        let link: Hyperlink = quick_xml::de::from_str(xml).unwrap();
        // Empty rid should result in default External with empty rid
        assert!(matches!(
            link.link,
            HyperlinkData::External { ref rid, .. } if rid.is_empty()
        ));
    }
}
