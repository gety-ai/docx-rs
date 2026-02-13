use crate::documents::{BuildXML, Level, LevelJc, LevelText, NumberFormat, Start};
use crate::types::LevelSuffixType;
use crate::xml_builder::*;
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;
use std::str::FromStr;

use super::style::{
    parse_paragraph_property_xml, parse_run_property_xml, ParagraphPropertyXml, RunPropertyXml,
    XmlValueAttr,
};

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AbstractNumbering {
    pub id: usize,
    pub style_link: Option<String>,
    pub num_style_link: Option<String>,
    pub levels: Vec<Level>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub multi_level_type: Option<String>,
}

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct LevelXml {
    #[serde(rename = "@ilvl", alias = "@w:ilvl", default)]
    pub level: Option<String>,
    #[serde(rename = "start", alias = "w:start", default)]
    pub start: Option<XmlValueAttr>,
    #[serde(rename = "numFmt", alias = "w:numFmt", default)]
    pub number_format: Option<XmlValueAttr>,
    #[serde(rename = "lvlText", alias = "w:lvlText", default)]
    pub level_text: Option<XmlValueAttr>,
    #[serde(rename = "lvlJc", alias = "w:lvlJc", default)]
    pub level_jc: Option<XmlValueAttr>,
    #[serde(rename = "pPr", alias = "w:pPr", default)]
    pub paragraph_property: Option<ParagraphPropertyXml>,
    #[serde(rename = "rPr", alias = "w:rPr", default)]
    pub run_property: Option<RunPropertyXml>,
    #[serde(rename = "suff", alias = "w:suff", default)]
    pub suffix: Option<XmlValueAttr>,
    #[serde(rename = "pStyle", alias = "w:pStyle", default)]
    pub paragraph_style: Option<XmlValueAttr>,
    #[serde(rename = "lvlRestart", alias = "w:lvlRestart", default)]
    pub level_restart: Option<XmlValueAttr>,
    #[serde(rename = "isLgl", alias = "w:isLgl", default)]
    pub is_lgl: Option<XmlValueAttr>,
}

#[derive(Debug, Deserialize, Default)]
struct AbstractNumberingXml {
    #[serde(rename = "@abstractNumId", alias = "@w:abstractNumId", default)]
    id: Option<String>,
    #[serde(rename = "styleLink", alias = "w:styleLink", default)]
    style_link: Option<XmlValueAttr>,
    #[serde(rename = "numStyleLink", alias = "w:numStyleLink", default)]
    num_style_link: Option<XmlValueAttr>,
    #[serde(rename = "multiLevelType", alias = "w:multiLevelType", default)]
    multi_level_type: Option<XmlValueAttr>,
    #[serde(rename = "lvl", alias = "w:lvl", default)]
    levels: Vec<LevelXml>,
}

pub(crate) fn parse_usize_attr(value: Option<String>, default: usize) -> usize {
    value
        .and_then(|v| {
            v.parse::<usize>()
                .ok()
                .or_else(|| v.parse::<f32>().ok().map(|f| f as usize))
        })
        .unwrap_or(default)
}

pub(crate) fn level_from_xml(xml: LevelXml) -> Level {
    let level = parse_usize_attr(xml.level, 0);
    let start = parse_usize_attr(xml.start.and_then(|v| v.val), 1);
    let number_format = xml
        .number_format
        .and_then(|v| v.val)
        .unwrap_or_else(|| "decimal".to_string());
    let level_text = xml.level_text.and_then(|v| v.val).unwrap_or_default();
    let level_jc = xml
        .level_jc
        .and_then(|v| v.val)
        .unwrap_or_else(|| "left".to_string());

    let mut out = Level::new(
        level,
        Start::new(start),
        NumberFormat::new(number_format),
        LevelText::new(level_text),
        LevelJc::new(level_jc),
    );

    if let Some(v) = xml.paragraph_style.and_then(|v| v.val) {
        out = out.paragraph_style(v);
    }
    if let Some(v) = xml.suffix.and_then(|v| v.val) {
        if let Ok(suffix) = LevelSuffixType::from_str(&v) {
            out = out.suffix(suffix);
        }
    }
    if let Some(v) = xml.level_restart.and_then(|v| v.val) {
        if let Ok(n) = v.parse::<u32>() {
            out = out.level_restart(n);
        }
    }
    if xml.is_lgl.is_some() {
        out = out.is_lgl();
    }

    out.paragraph_property = parse_paragraph_property_xml(xml.paragraph_property);
    out.run_property = parse_run_property_xml(xml.run_property);
    out
}

impl<'de> Deserialize<'de> for AbstractNumbering {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = AbstractNumberingXml::deserialize(deserializer)?;
        let mut abs = AbstractNumbering::new(parse_usize_attr(xml.id, 0));
        abs.multi_level_type = xml.multi_level_type.and_then(|v| v.val);
        if let Some(v) = xml.style_link.and_then(|v| v.val) {
            abs = abs.style_link(v);
        }
        if let Some(v) = xml.num_style_link.and_then(|v| v.val) {
            abs = abs.num_style_link(v);
        }
        for level in xml.levels {
            abs = abs.add_level(level_from_xml(level));
        }
        Ok(abs)
    }
}

impl AbstractNumbering {
    pub fn new(id: usize) -> Self {
        Self {
            id,
            style_link: None,
            num_style_link: None,
            levels: vec![],
            multi_level_type: None,
        }
    }

    pub fn add_level(mut self, level: Level) -> Self {
        self.levels.push(level);
        self
    }

    pub fn num_style_link(mut self, link: impl Into<String>) -> Self {
        self.num_style_link = Some(link.into());
        self
    }

    pub fn style_link(mut self, link: impl Into<String>) -> Self {
        self.style_link = Some(link.into());
        self
    }
}

impl BuildXML for AbstractNumbering {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        let mut builder = XMLBuilder::from(stream)  
            .open_abstract_num(&self.id.to_string())?;  
          
        // 添加 multiLevelType 元素（如果存在）  
        if let Some(ref multi_level_type) = self.multi_level_type {  
            builder = builder.multi_level_type(multi_level_type)?;  
        }  
          
        builder  
            .add_children(&self.levels)?  
            .close()?  
            .into_inner()  
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use crate::documents::{Level, LevelJc, LevelText, NumberFormat, Start};
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_numbering() {
        let mut c = AbstractNumbering::new(0);
        c = c.add_level(Level::new(
            1,
            Start::new(1),
            NumberFormat::new("decimal"),
            LevelText::new("%4."),
            LevelJc::new("left"),
        ));
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="1"><w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%4." /><w:lvlJc w:val="left" /><w:pPr><w:rPr /></w:pPr><w:rPr /></w:lvl></w:abstractNum>"#
        );
    }

    #[test]
    fn test_numbering_json() {
        let mut c = AbstractNumbering::new(0);
        c = c
            .add_level(Level::new(
                1,
                Start::new(1),
                NumberFormat::new("decimal"),
                LevelText::new("%4."),
                LevelJc::new("left"),
            ))
            .num_style_link("style1");
        assert_eq!(
            serde_json::to_string(&c).unwrap(),
            r#"{"id":0,"styleLink":null,"numStyleLink":"style1","levels":[{"level":1,"start":1,"format":"decimal","text":"%4.","jc":"left","paragraphProperty":{"runProperty":{},"tabs":[]},"runProperty":{},"suffix":"tab","pstyle":null,"levelRestart":null}]}"#,
        );
    }
}
