use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use crate::documents::BuildXML;
use crate::xml_builder::*;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CommentExtended {
    #[serde(
        rename(serialize = "paragraphId", deserialize = "@paraId"),
        alias = "@w15:paraId",
        alias = "paragraphId"
    )]
    pub paragraph_id: String,
    #[serde(
        rename(serialize = "done", deserialize = "@done"),
        alias = "@w15:done",
        alias = "done",
        default,
        deserialize_with = "deserialize_done"
    )]
    pub done: bool,
    #[serde(
        rename(serialize = "parentParagraphId", deserialize = "@paraIdParent"),
        alias = "@w15:paraIdParent",
        alias = "parentParagraphId",
        default
    )]
    pub parent_paragraph_id: Option<String>,
}

impl CommentExtended {
    pub fn new(paragraph_id: impl Into<String>) -> CommentExtended {
        Self {
            paragraph_id: paragraph_id.into(),
            done: false,
            parent_paragraph_id: None,
        }
    }

    pub fn done(mut self) -> CommentExtended {
        self.done = true;
        self
    }

    pub fn parent_paragraph_id(mut self, id: impl Into<String>) -> CommentExtended {
        self.parent_paragraph_id = Some(id.into());
        self
    }
}

fn deserialize_done<'de, D>(deserializer: D) -> Result<bool, D::Error>
where
    D: Deserializer<'de>,
{
    #[derive(Deserialize)]
    #[serde(untagged)]
    enum DoneValue {
        Bool(bool),
        String(String),
        Number(u8),
    }

    match DoneValue::deserialize(deserializer)? {
        DoneValue::Bool(v) => Ok(v),
        DoneValue::Number(v) => Ok(v != 0),
        DoneValue::String(v) => {
            let normalized = v.trim().to_ascii_lowercase();
            match normalized.as_str() {
                "1" | "true" => Ok(true),
                "0" | "false" | "" => Ok(false),
                _ => Err(serde::de::Error::custom(
                    "invalid done value, expected 0/1/true/false",
                )),
            }
        }
    }
}

impl BuildXML for CommentExtended {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .comment_extended(&self.paragraph_id, self.done, &self.parent_paragraph_id)?
            .into_inner()
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    #[test]
    fn test_comment_extended_json() {
        let ex = CommentExtended {
            paragraph_id: "00002".to_owned(),
            done: false,
            parent_paragraph_id: Some("0004".to_owned()),
        };
        assert_eq!(
            serde_json::to_string(&ex).unwrap(),
            r#"{"paragraphId":"00002","done":false,"parentParagraphId":"0004"}"#
        );
    }
}
