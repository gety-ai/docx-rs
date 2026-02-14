use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use super::*;

use crate::documents::{BuildXML, HistoryId, Run};
use crate::{escape, xml_builder::*};

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct InsertXml {
    #[serde(rename = "@id", alias = "@w:id", default)]
    _id: Option<String>,
    #[serde(rename = "@author", alias = "@w:author", default)]
    author: Option<String>,
    #[serde(rename = "@date", alias = "@w:date", default)]
    date: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<InsertChildXml>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlIdNode {
    #[serde(rename = "@id", alias = "@w:id", default)]
    id: Option<String>,
}

#[derive(Debug, Deserialize)]
enum InsertChildXml {
    #[serde(rename = "r", alias = "w:r")]
    Run(Run),
    #[serde(rename = "del", alias = "w:del")]
    Delete(Delete),
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

fn insert_child_from_xml(xml: InsertChildXml) -> Option<InsertChild> {
    match xml {
        InsertChildXml::Run(run) => Some(InsertChild::Run(Box::new(run))),
        InsertChildXml::Delete(delete) => Some(InsertChild::Delete(delete)),
        InsertChildXml::CommentStart(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(InsertChild::CommentStart(Box::new(CommentRangeStart::new(
                Comment::new(id),
            ))))
        }
        InsertChildXml::CommentEnd(node) => {
            let id = parse_optional_usize(node.id)?;
            Some(InsertChild::CommentEnd(CommentRangeEnd::new(id)))
        }
        InsertChildXml::Unknown => None,
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum InsertChild {
    Run(Box<Run>),
    Delete(Delete),
    CommentStart(Box<CommentRangeStart>),
    CommentEnd(CommentRangeEnd),
}

impl BuildXML for InsertChild {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        match self {
            InsertChild::Run(v) => v.build_to(stream),
            InsertChild::Delete(v) => v.build_to(stream),
            InsertChild::CommentStart(v) => v.build_to(stream),
            InsertChild::CommentEnd(v) => v.build_to(stream),
        }
    }
}

impl Serialize for InsertChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            InsertChild::Run(ref r) => {
                let mut t = serializer.serialize_struct("Run", 2)?;
                t.serialize_field("type", "run")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            InsertChild::Delete(ref r) => {
                let mut t = serializer.serialize_struct("Delete", 2)?;
                t.serialize_field("type", "delete")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            InsertChild::CommentStart(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeStart", 2)?;
                t.serialize_field("type", "commentRangeStart")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            InsertChild::CommentEnd(ref r) => {
                let mut t = serializer.serialize_struct("CommentRangeEnd", 2)?;
                t.serialize_field("type", "commentRangeEnd")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct Insert {
    pub children: Vec<InsertChild>,
    pub author: String,
    pub date: String,
}

impl<'de> Deserialize<'de> for Insert {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = InsertXml::deserialize(deserializer)?;
        let mut insert = Insert::default();

        if let Some(author) = xml.author {
            insert.author = author;
        }
        if let Some(date) = xml.date {
            insert.date = date;
        }

        insert.children = xml
            .children
            .into_iter()
            .filter_map(insert_child_from_xml)
            .collect();
        Ok(insert)
    }
}

impl Default for Insert {
    fn default() -> Insert {
        Insert {
            author: "unnamed".to_owned(),
            date: "1970-01-01T00:00:00Z".to_owned(),
            children: vec![],
        }
    }
}

impl Insert {
    pub fn new(run: Run) -> Insert {
        Self {
            children: vec![InsertChild::Run(Box::new(run))],
            ..Default::default()
        }
    }

    pub fn new_with_empty() -> Insert {
        Self {
            ..Default::default()
        }
    }

    pub fn new_with_del(del: Delete) -> Insert {
        Self {
            children: vec![InsertChild::Delete(del)],
            ..Default::default()
        }
    }

    pub fn add_run(mut self, run: Run) -> Insert {
        self.children.push(InsertChild::Run(Box::new(run)));
        self
    }

    pub fn add_delete(mut self, del: Delete) -> Insert {
        self.children.push(InsertChild::Delete(del));
        self
    }

    pub fn add_child(mut self, c: InsertChild) -> Insert {
        self.children.push(c);
        self
    }

    pub fn add_comment_start(mut self, comment: Comment) -> Self {
        self.children
            .push(InsertChild::CommentStart(Box::new(CommentRangeStart::new(
                comment,
            ))));
        self
    }

    pub fn add_comment_end(mut self, id: usize) -> Self {
        self.children
            .push(InsertChild::CommentEnd(CommentRangeEnd::new(id)));
        self
    }

    pub fn author(mut self, author: impl Into<String>) -> Insert {
        self.author = escape::escape(&author.into());
        self
    }

    pub fn date(mut self, date: impl Into<String>) -> Insert {
        self.date = date.into();
        self
    }
}

impl HistoryId for Insert {}

impl BuildXML for Insert {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_insert(&self.generate(), &self.author, &self.date)?
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
    fn test_ins_default() {
        let b = Insert::new(Run::new()).build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:ins w:id="123" w:author="unnamed" w:date="1970-01-01T00:00:00Z"><w:r><w:rPr /></w:r></w:ins>"#
        );
    }

    #[test]
    fn test_insert_xml_deserialize() {
        let xml = r#"<w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="0" w:author="John" w:date="2024-01-01T00:00:00Z">
            <w:r><w:t>inserted text</w:t></w:r>
            <w:commentRangeStart w:id="5"/>
            <w:commentRangeEnd w:id="5"/>
        </w:ins>"#;

        let ins: Insert = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(ins.author, "John");
        assert_eq!(ins.date, "2024-01-01T00:00:00Z");
        assert_eq!(ins.children.len(), 3);
        assert!(matches!(&ins.children[0], InsertChild::Run(_)));
        assert!(matches!(
            &ins.children[1],
            InsertChild::CommentStart(c) if c.id == 5
        ));
        assert!(matches!(
            &ins.children[2],
            InsertChild::CommentEnd(c) if c == &CommentRangeEnd::new(5)
        ));
    }
}
