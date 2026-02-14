use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use crate::xml_builder::*;
use crate::{documents::*, escape};

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct DeleteXml {
    #[serde(rename = "@id", alias = "@w:id", default)]
    _id: Option<String>,
    #[serde(rename = "@author", alias = "@w:author", default)]
    author: Option<String>,
    #[serde(rename = "@date", alias = "@w:date", default)]
    date: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<DeleteChildXml>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlIdNode {
    #[serde(rename = "@id", alias = "@w:id", default)]
    id: Option<String>,
}

#[derive(Debug, Deserialize)]
enum DeleteChildXml {
    #[serde(rename = "r", alias = "w:r")]
    Run(Run),
    #[serde(rename = "commentRangeStart", alias = "w:commentRangeStart")]
    CommentStart(XmlIdNode),
    #[serde(rename = "commentRangeEnd", alias = "w:commentRangeEnd")]
    CommentEnd(XmlIdNode),
    #[serde(other)]
    Unknown,
}

fn parse_optional_usize(v: Option<String>) -> Option<usize> {
    v.and_then(|s| s.parse::<usize>().ok())
}

fn delete_child_from_xml(xml: DeleteChildXml) -> Option<DeleteChild> {
    match xml {
        DeleteChildXml::Run(run) => Some(DeleteChild::Run(run)),
        DeleteChildXml::CommentStart(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(DeleteChild::CommentStart(Box::new(CommentRangeStart::new(
                Comment::new(id),
            ))))
        }
        DeleteChildXml::CommentEnd(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(DeleteChild::CommentEnd(CommentRangeEnd::new(id)))
        }
        DeleteChildXml::Unknown => None,
    }
}

#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct Delete {
    pub author: String,
    pub date: String,
    pub children: Vec<DeleteChild>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum DeleteChild {
    Run(Run),
    CommentStart(Box<CommentRangeStart>),
    CommentEnd(CommentRangeEnd),
}

impl<'de> Deserialize<'de> for Delete {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = DeleteXml::deserialize(deserializer)?;
        let mut delete = Delete::default();

        if let Some(author) = xml.author {
            delete.author = author;
        }
        if let Some(date) = xml.date {
            delete.date = date;
        }

        delete.children = xml
            .children
            .into_iter()
            .filter_map(delete_child_from_xml)
            .collect();
        Ok(delete)
    }
}

impl Serialize for DeleteChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            DeleteChild::Run(ref r) => {
                let mut t = serializer.serialize_struct("Run", 2)?;
                t.serialize_field("type", "run")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            DeleteChild::CommentStart(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeStart", 2)?;
                t.serialize_field("type", "commentRangeStart")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            DeleteChild::CommentEnd(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeEnd", 2)?;
                t.serialize_field("type", "commentRangeEnd")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

impl Default for Delete {
    fn default() -> Delete {
        Delete {
            author: "unnamed".to_owned(),
            date: "1970-01-01T00:00:00Z".to_owned(),
            children: vec![],
        }
    }
}

impl Delete {
    pub fn new() -> Delete {
        Self {
            children: vec![],
            ..Default::default()
        }
    }

    pub fn add_run(mut self, run: Run) -> Delete {
        self.children.push(DeleteChild::Run(run));
        self
    }

    pub fn add_comment_start(mut self, comment: Comment) -> Delete {
        self.children
            .push(DeleteChild::CommentStart(Box::new(CommentRangeStart::new(
                comment,
            ))));
        self
    }

    pub fn add_comment_end(mut self, id: usize) -> Delete {
        self.children
            .push(DeleteChild::CommentEnd(CommentRangeEnd::new(id)));
        self
    }

    pub fn author(mut self, author: impl Into<String>) -> Delete {
        self.author = escape::escape(&author.into());
        self
    }

    pub fn date(mut self, date: impl Into<String>) -> Delete {
        self.date = date.into();
        self
    }
}

impl HistoryId for Delete {}

impl BuildXML for Delete {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        let id = self.generate();
        XMLBuilder::from(stream)
            .open_delete(&id, &self.author, &self.date)?
            .apply_each(&self.children, |ch, b| match ch {
                DeleteChild::Run(t) => b.add_child(t),
                DeleteChild::CommentStart(c) => b.add_child(&c),
                DeleteChild::CommentEnd(c) => b.add_child(c),
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
    fn test_delete_default() {
        let b = Delete::new().add_run(Run::new()).build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:del w:id="123" w:author="unnamed" w:date="1970-01-01T00:00:00Z"><w:r><w:rPr /></w:r></w:del>"#
        );
    }

    #[test]
    fn test_delete_xml_deserialize() {
        let xml = r#"<w:del xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="3" w:author="Jane" w:date="2024-01-03T00:00:00Z">
            <w:r><w:delText>deleted text</w:delText></w:r>
            <w:commentRangeStart w:id="6"/>
            <w:commentRangeEnd w:id="6"/>
        </w:del>"#;

        let del: Delete = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(del.author, "Jane");
        assert_eq!(del.date, "2024-01-03T00:00:00Z");
        assert_eq!(del.children.len(), 3);
        assert!(matches!(&del.children[0], DeleteChild::Run(_)));
        assert!(matches!(
            &del.children[1],
            DeleteChild::CommentStart(c) if c.id == 6
        ));
        assert!(matches!(
            &del.children[2],
            DeleteChild::CommentEnd(c) if c == &CommentRangeEnd::new(6)
        ));
    }
}
