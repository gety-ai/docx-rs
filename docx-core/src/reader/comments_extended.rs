use quick_xml::de::from_reader;
use std::io::{BufReader, Read};

use super::*;

fn dedup_by_paragraph_id(children: Vec<CommentExtended>) -> Vec<CommentExtended> {
    let mut deduped: Vec<CommentExtended> = Vec::with_capacity(children.len());
    for ex in children {
        if let Some(pos) = deduped
            .iter()
            .position(|current: &CommentExtended| current.paragraph_id == ex.paragraph_id)
        {
            deduped[pos] = ex;
        } else {
            deduped.push(ex);
        }
    }
    deduped
}

impl FromXMLQuickXml for CommentsExtended {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        let parsed: CommentsExtended = from_reader(BufReader::new(reader))?;
        Ok(CommentsExtended {
            children: dedup_by_paragraph_id(parsed.children),
        })
    }
}

impl FromXML for CommentsExtended {
    fn from_xml<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Self::from_xml_quick(reader)
    }
}
