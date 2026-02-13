use crate::documents::BuildXML;
use crate::xml_builder::*;
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use super::abstract_numbering::{level_from_xml, parse_usize_attr, LevelXml};
use super::style::XmlValueAttr;
use super::*;

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Numbering {
    pub id: usize,
    pub abstract_num_id: usize,
    pub level_overrides: Vec<LevelOverride>,
}

#[derive(Debug, Deserialize, Default)]
struct LevelOverrideXml {
    #[serde(rename = "@ilvl", alias = "@w:ilvl", default)]
    level: Option<String>,
    #[serde(rename = "startOverride", alias = "w:startOverride", default)]
    override_start: Option<XmlValueAttr>,
    #[serde(rename = "lvl", alias = "w:lvl", default)]
    override_level: Option<LevelXml>,
}

#[derive(Debug, Deserialize, Default)]
struct NumberingXml {
    #[serde(rename = "@numId", alias = "@w:numId", default)]
    id: Option<String>,
    #[serde(rename = "abstractNumId", alias = "w:abstractNumId", default)]
    abstract_num_id: Option<XmlValueAttr>,
    #[serde(rename = "lvlOverride", alias = "w:lvlOverride", default)]
    level_overrides: Vec<LevelOverrideXml>,
}

impl<'de> Deserialize<'de> for Numbering {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = NumberingXml::deserialize(deserializer)?;
        let id = parse_usize_attr(xml.id, 0);
        let abstract_num_id = parse_usize_attr(xml.abstract_num_id.and_then(|v| v.val), 0);
        let mut numbering = Numbering::new(id, abstract_num_id);
        for item in xml.level_overrides {
            let mut o = LevelOverride::new(parse_usize_attr(item.level, 0));
            if let Some(v) = item.override_start.and_then(|v| v.val) {
                if let Ok(n) = v.parse::<usize>() {
                    o = o.start(n);
                }
            }
            if let Some(lvl) = item.override_level {
                o = o.level(level_from_xml(lvl));
            }
            numbering = numbering.add_override(o);
        }
        Ok(numbering)
    }
}

impl Numbering {
    pub fn new(id: usize, abstract_num_id: usize) -> Self {
        Self {
            id,
            abstract_num_id,
            level_overrides: vec![],
        }
    }

    pub fn overrides(mut self, overrides: Vec<LevelOverride>) -> Self {
        self.level_overrides = overrides;
        self
    }

    pub fn add_override(mut self, o: LevelOverride) -> Self {
        self.level_overrides.push(o);
        self
    }
}

impl BuildXML for Numbering {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        let id = format!("{}", self.id);
        let abs_id = format!("{}", self.abstract_num_id);
        XMLBuilder::from(stream)
            .open_num(&id)?
            .abstract_num_id(&abs_id)?
            .add_children(&self.level_overrides)?
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
    fn test_numbering() {
        let c = Numbering::new(0, 2);
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:num w:numId="0"><w:abstractNumId w:val="2" /></w:num>"#
        );
    }
    #[test]
    fn test_numbering_override() {
        let c = Numbering::new(0, 2);
        let overrides = vec![
            LevelOverride::new(0).start(1),
            LevelOverride::new(1).start(1),
        ];
        let b = c.overrides(overrides).build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:num w:numId="0"><w:abstractNumId w:val="2" /><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1" /></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:startOverride w:val="1" /></w:lvlOverride></w:num>"#
        );
    }

    #[test]
    fn test_numbering_override_json() {
        let c = Numbering::new(0, 2);
        let overrides = vec![
            LevelOverride::new(0).start(1),
            LevelOverride::new(1).start(1),
        ];
        assert_eq!(
            serde_json::to_string(&c.overrides(overrides)).unwrap(),
            r#"{"id":0,"abstractNumId":2,"levelOverrides":[{"level":0,"overrideStart":1,"overrideLevel":null},{"level":1,"overrideStart":1,"overrideLevel":null}]}"#
        );
    }
}
