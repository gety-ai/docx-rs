use super::*;
use std::io::Write;
use std::str::FromStr;

use crate::documents::BuildXML;
use crate::types::CharacterSpacingValues;
use crate::xml_builder::*;

use serde::{Deserialize, Deserializer, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Settings {
    default_tab_stop: DefaultTabStop,
    zoom: Zoom,
    doc_id: Option<DocId>,
    doc_vars: Vec<DocVar>,
    even_and_odd_headers: bool,
    adjust_line_height_in_table: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    character_spacing_control: Option<CharacterSpacingValues>,
}

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct SettingsXml {
    #[serde(rename = "$value", default)]
    children: Vec<SettingsChildXml>,
}

#[derive(Debug, Deserialize)]
enum SettingsChildXml {
    #[serde(rename = "defaultTabStop", alias = "w:defaultTabStop")]
    DefaultTabStop(SettingsDefaultTabStopXml),
    #[serde(rename = "zoom", alias = "w:zoom")]
    Zoom(SettingsZoomXml),
    #[serde(rename = "docId", alias = "w:docId", alias = "w14:docId", alias = "w15:docId")]
    DocId(SettingsDocIdXml),
    #[serde(rename = "docVars", alias = "w:docVars")]
    DocVars(SettingsDocVarsXml),
    #[serde(rename = "docVar", alias = "w:docVar")]
    DocVar(SettingsDocVarXml),
    #[serde(rename = "evenAndOddHeaders", alias = "w:evenAndOddHeaders")]
    EvenAndOddHeaders(SettingsOnOffXml),
    #[serde(rename = "adjustLineHeightInTable", alias = "w:adjustLineHeightInTable")]
    AdjustLineHeightInTable(SettingsOnOffXml),
    #[serde(rename = "characterSpacingControl", alias = "w:characterSpacingControl")]
    CharacterSpacingControl(SettingsValueXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsDefaultTabStopXml {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsZoomXml {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
    #[serde(rename = "@percent", alias = "@w:percent", default)]
    percent: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsDocIdXml {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
    #[serde(rename = "@w14:val", default)]
    w14_val: Option<String>,
    #[serde(rename = "@w15:val", default)]
    w15_val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsDocVarsXml {
    #[serde(rename = "docVar", alias = "w:docVar", default)]
    doc_vars: Vec<SettingsDocVarXml>,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsDocVarXml {
    #[serde(rename = "@name", alias = "@w:name", default)]
    name: Option<String>,
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsOnOffXml {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SettingsValueXml {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

fn normalize_doc_id(raw: Option<String>) -> Option<String> {
    raw.map(|v| v.replace(['{', '}'], "")).and_then(|v| {
        let trimmed = v.trim();
        if trimmed.is_empty() {
            None
        } else {
            Some(trimmed.to_owned())
        }
    })
}

impl<'de> Deserialize<'de> for Settings {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = SettingsXml::deserialize(deserializer)?;
        let mut settings = Settings::default();

        let mut doc_vars_nested = Vec::new();
        let mut doc_vars_direct = Vec::new();
        let mut doc_ids: Vec<SettingsDocIdXml> = Vec::new();

        for child in xml.children {
            match child {
                SettingsChildXml::DefaultTabStop(node) => {
                    if let Some(val) =
                        node.val.and_then(|v| v.parse::<f32>().ok().map(|v| v as usize))
                    {
                        settings.default_tab_stop = DefaultTabStop::new(val);
                    }
                }
                SettingsChildXml::Zoom(node) => {
                    // Try percent first, then val
                    let value = node.percent.or(node.val);
                    if let Some(val) = value.and_then(|v| v.parse::<usize>().ok()) {
                        settings.zoom = Zoom::new(val);
                    }
                }
                SettingsChildXml::DocId(node) => {
                    doc_ids.push(node);
                }
                SettingsChildXml::DocVars(node) => {
                    doc_vars_nested.extend(node.doc_vars);
                }
                SettingsChildXml::DocVar(node) => {
                    doc_vars_direct.push(node);
                }
                SettingsChildXml::EvenAndOddHeaders(node) => {
                    settings.even_and_odd_headers = node
                        .val
                        .map(|v| {
                            let normalized = v.trim().to_ascii_lowercase();
                            normalized != "0" && normalized != "false"
                        })
                        .unwrap_or(true);
                }
                SettingsChildXml::AdjustLineHeightInTable(node) => {
                    settings.adjust_line_height_in_table = node
                        .val
                        .map(|v| {
                            let normalized = v.trim().to_ascii_lowercase();
                            normalized != "0" && normalized != "false"
                        })
                        .unwrap_or(true);
                }
                SettingsChildXml::CharacterSpacingControl(node) => {
                    if let Some(val) = node
                        .val
                        .and_then(|v| CharacterSpacingValues::from_str(&v).ok())
                    {
                        settings.character_spacing_control = Some(val);
                    }
                }
                SettingsChildXml::Unknown => {}
            }
        }

        settings.doc_vars = doc_vars_nested
            .into_iter()
            .chain(doc_vars_direct.into_iter())
            .filter_map(|var| match (var.name, var.val) {
                (Some(name), Some(val)) => Some(DocVar::new(name, val)),
                _ => None,
            })
            .collect();

        // DocId priority: w15:val > val > w14:val (last occurrence wins within same priority)
        let doc_id_w15 = doc_ids
            .iter()
            .rev()
            .find_map(|doc_id| normalize_doc_id(doc_id.w15_val.clone()));
        let doc_id_default = doc_ids
            .iter()
            .rev()
            .find_map(|doc_id| normalize_doc_id(doc_id.val.clone()));
        let doc_id_w14 = doc_ids
            .iter()
            .rev()
            .find_map(|doc_id| normalize_doc_id(doc_id.w14_val.clone()));

        settings.doc_id = doc_id_w15.or(doc_id_default).or(doc_id_w14).map(DocId::new);

        Ok(settings)
    }
}

impl Settings {
    pub fn new() -> Settings {
        Default::default()
    }

    pub fn doc_id(mut self, id: impl Into<String>) -> Self {
        self.doc_id = Some(DocId::new(id.into()));
        self
    }

    pub fn default_tab_stop(mut self, tab_stop: usize) -> Self {
        self.default_tab_stop = DefaultTabStop::new(tab_stop);
        self
    }

    pub fn add_doc_var(mut self, name: impl Into<String>, val: impl Into<String>) -> Self {
        self.doc_vars.push(DocVar::new(name, val));
        self
    }

    pub fn even_and_odd_headers(mut self) -> Self {
        self.even_and_odd_headers = true;
        self
    }

    pub fn adjust_line_height_in_table(mut self) -> Self {
        self.adjust_line_height_in_table = true;
        self
    }

    pub fn character_spacing_control(mut self, val: CharacterSpacingValues) -> Self {
        self.character_spacing_control = Some(val);
        self
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            default_tab_stop: DefaultTabStop::new(840),
            zoom: Zoom::new(100),
            doc_id: None,
            doc_vars: vec![],
            even_and_odd_headers: false,
            adjust_line_height_in_table: false,
            character_spacing_control: None,
        }
    }
}

impl BuildXML for Settings {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .declaration(Some(true))?
            .open_settings()?
            .add_child(&self.default_tab_stop)?
            .add_child(&self.zoom)?
            .open_compat()?
            .space_for_ul()?
            .balance_single_byte_double_byte_width()?
            .do_not_leave_backslash_alone()?
            .ul_trail_space()?
            .do_not_expand_shift_return()?
            .apply_opt(self.character_spacing_control, |v, b| {
                b.character_spacing_control(&v.to_string())
            })?
            .apply_if(self.adjust_line_height_in_table, |b| {
                b.adjust_line_height_table()
            })?
            .use_fe_layout()?
            .compat_setting(
                "compatibilityMode",
                "http://schemas.microsoft.com/office/word",
                "15",
            )?
            .compat_setting(
                "overrideTableStyleFontSizeAndJustification",
                "http://schemas.microsoft.com/office/word",
                "1",
            )?
            .compat_setting(
                "enableOpenTypeFeatures",
                "http://schemas.microsoft.com/office/word",
                "1",
            )?
            .compat_setting(
                "doNotFlipMirrorIndents",
                "http://schemas.microsoft.com/office/word",
                "1",
            )?
            .compat_setting(
                "differentiateMultirowTableHeaders",
                "http://schemas.microsoft.com/office/word",
                "1",
            )?
            .compat_setting(
                "useWord2013TrackBottomHyphenation",
                "http://schemas.microsoft.com/office/word",
                "0",
            )?
            .close()?
            .add_optional_child(&self.doc_id)?
            .apply_if(!self.doc_vars.is_empty(), |b| {
                b.open_doc_vars()?.add_children(&self.doc_vars)?.close()
            })?
            .apply_if(self.even_and_odd_headers, |b| b.even_and_odd_headers())?
            .close()?
            .into_inner()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use quick_xml::de::from_str;
    use std::str;

    #[test]
    fn test_settings() {
        let c = Settings::new();
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:defaultTabStop w:val="840" /><w:zoom w:percent="100" /><w:compat><w:spaceForUL /><w:balanceSingleByteDoubleByteWidth /><w:doNotLeaveBackslashAlone /><w:ulTrailSpace /><w:doNotExpandShiftReturn /><w:useFELayout /><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15" /><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1" /><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1" /><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1" /><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1" /><w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0" /></w:compat></w:settings>"#
        );
    }

    #[test]
    fn test_settings_deserialize_prefers_w15_doc_id() {
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w14:docId w14:val="4A6FEB4F"/><w15:docId w15:val="{C11ED300-8EA6-3D41-8D67-5E5DE3410CF8}"/></w:settings>"#;
        let settings: Settings = from_str(xml).unwrap();
        assert_eq!(
            settings.doc_id,
            Some(DocId::new("C11ED300-8EA6-3D41-8D67-5E5DE3410CF8"))
        );
    }

    #[test]
    fn test_settings_deserialize_w14_doc_id_fallback() {
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:docId w14:val="4A6FEB4F"/></w:settings>"#;
        let settings: Settings = from_str(xml).unwrap();
        assert_eq!(settings.doc_id, Some(DocId::new("4A6FEB4F")));
    }

    #[test]
    fn test_settings_deserialize_basic_flags_and_doc_vars() {
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/><w:zoom w:percent="125"/><w:docVars><w:docVar w:name="foo" w:val="bar"/></w:docVars><w:docVar w:name="baz" w:val="qux"/><w:evenAndOddHeaders w:val="0"/><w:adjustLineHeightInTable/><w:characterSpacingControl w:val="compressPunctuation"/></w:settings>"#;
        let settings: Settings = from_str(xml).unwrap();

        assert_eq!(settings.default_tab_stop, DefaultTabStop::new(720));
        assert_eq!(settings.zoom, Zoom::new(125));
        assert_eq!(
            settings.doc_vars,
            vec![DocVar::new("foo", "bar"), DocVar::new("baz", "qux")]
        );
        assert!(!settings.even_and_odd_headers);
        assert!(settings.adjust_line_height_in_table);
        assert_eq!(
            settings.character_spacing_control,
            Some(CharacterSpacingValues::CompressPunctuation)
        );
    }
}
