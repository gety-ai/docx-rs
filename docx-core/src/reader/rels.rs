use super::*;
use crate::reader::{FromXML, FromXMLQuickXml, ReaderError};
use serde::Deserialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    io::{BufReader, Read},
    path::{Path, PathBuf},
};

pub type ReadRels = BTreeMap<String, BTreeSet<(RId, PathBuf, Option<String>)>>;

// ============================================================================
// XML Deserialization DTOs (quick-xml serde)
// ============================================================================

#[derive(Deserialize)]
struct RelationshipXml {
    #[serde(rename = "@Type", default)]
    rel_type: String,
    #[serde(rename = "@Id", default)]
    id: String,
    #[serde(rename = "@Target", default)]
    target: String,
    #[serde(rename = "@TargetMode", default)]
    target_mode: Option<String>,
}

#[derive(Deserialize)]
enum RelationshipsChildXml {
    Relationship(RelationshipXml),
    #[serde(other)]
    Unknown,
}

#[derive(Deserialize)]
struct RelationshipsXml {
    #[serde(rename = "$value", default)]
    children: Vec<RelationshipsChildXml>,
}

impl FromXMLQuickXml for Rels {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        let xml: RelationshipsXml = quick_xml::de::from_reader(BufReader::new(reader))?;
        let mut s = Self::default();
        for child in xml.children {
            if let RelationshipsChildXml::Relationship(r) = child {
                s.rels.push((r.rel_type, r.id, r.target));
            }
        }
        Ok(s)
    }
}

impl FromXML for Rels {
    fn from_xml<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Self::from_xml_quick(reader)
    }
}

pub fn find_rels_filename(main_path: impl AsRef<Path>) -> Result<PathBuf, ReaderError> {
    let path = main_path.as_ref();
    let dir = path
        .parent()
        .ok_or(ReaderError::DocumentRelsNotFoundError)?;
    let base = path
        .file_stem()
        .ok_or(ReaderError::DocumentRelsNotFoundError)?;
    Ok(Path::new(dir)
        .join("_rels")
        .join(base)
        .with_extension("xml.rels"))
}

pub fn read_rels_xml<R: Read>(reader: R, dir: impl AsRef<Path>) -> Result<ReadRels, ReaderError> {
    let xml: RelationshipsXml = quick_xml::de::from_reader(BufReader::new(reader))?;
    let mut rels: BTreeMap<String, BTreeSet<(RId, PathBuf, Option<String>)>> = BTreeMap::new();

    for child in xml.children {
        if let RelationshipsChildXml::Relationship(r) = child {
            let target = if !r.rel_type.ends_with("hyperlink") {
                Path::new(dir.as_ref()).join(&r.target)
            } else {
                Path::new("").join(&r.target)
            };

            rels.entry(r.rel_type)
                .or_insert_with(BTreeSet::new)
                .insert((r.id, target, r.target_mode));
        }
    }
    Ok(rels)
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;

    #[test]
    fn test_from_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
</Relationships>"#;
        let c = Rels::from_xml(xml.as_bytes()).unwrap();
        let rels =
            vec![
        (
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                .to_owned(),
            "rId1".to_owned(),
            "docProps/core.xml".to_owned(),
        )];
        assert_eq!(Rels { rels }, c);
    }
}
