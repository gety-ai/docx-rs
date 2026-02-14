use serde::{Deserialize, Deserializer, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize)]
#[cfg_attr(feature = "wasm", derive(ts_rs::TS))]
#[cfg_attr(feature = "wasm", ts(export))]
#[serde(rename_all = "camelCase")]
pub struct FontSchemeFont {
    pub script: String,
    pub typeface: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Default)]
#[cfg_attr(feature = "wasm", derive(ts_rs::TS))]
#[cfg_attr(feature = "wasm", ts(export))]
#[serde(rename_all = "camelCase")]
pub struct FontGroup {
    pub latin: String,
    pub ea: String,
    pub cs: String,
    pub fonts: Vec<FontSchemeFont>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Default)]
#[cfg_attr(feature = "wasm", derive(ts_rs::TS))]
#[cfg_attr(feature = "wasm", ts(export))]
#[serde(rename_all = "camelCase")]
pub struct FontScheme {
    pub major_font: FontGroup,
    pub minor_font: FontGroup,
}

// For now reader only
impl FontScheme {
    pub fn new() -> Self {
        Self::default()
    }
}

// ============================================================================
// XML Deserialization (quick-xml serde)
// ============================================================================

#[derive(Deserialize)]
struct FontTypefaceXml {
    #[serde(rename = "@typeface", default)]
    typeface: String,
}

#[derive(Deserialize)]
struct FontScriptXml {
    #[serde(rename = "@script", default)]
    script: String,
    #[serde(rename = "@typeface", default)]
    typeface: String,
}

#[derive(Deserialize)]
enum FontGroupChildXml {
    #[serde(rename = "latin", alias = "a:latin")]
    Latin(FontTypefaceXml),
    #[serde(rename = "ea", alias = "a:ea")]
    Ea(FontTypefaceXml),
    #[serde(rename = "cs", alias = "a:cs")]
    Cs(FontTypefaceXml),
    #[serde(rename = "font", alias = "a:font")]
    Font(FontScriptXml),
    #[serde(other)]
    Unknown,
}

#[derive(Deserialize)]
struct FontGroupXml {
    #[serde(rename = "$value", default)]
    children: Vec<FontGroupChildXml>,
}

impl<'de> Deserialize<'de> for FontGroup {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = FontGroupXml::deserialize(deserializer)?;
        let mut fg = FontGroup::default();
        for child in xml.children {
            match child {
                FontGroupChildXml::Latin(n) => fg.latin = n.typeface,
                FontGroupChildXml::Ea(n) => fg.ea = n.typeface,
                FontGroupChildXml::Cs(n) => fg.cs = n.typeface,
                FontGroupChildXml::Font(n) => {
                    fg.fonts.push(FontSchemeFont {
                        script: n.script,
                        typeface: n.typeface,
                    });
                }
                FontGroupChildXml::Unknown => {}
            }
        }
        Ok(fg)
    }
}

#[derive(Deserialize)]
struct FontSchemeXml {
    #[serde(rename = "majorFont", alias = "a:majorFont", default)]
    major_font: Option<FontGroup>,
    #[serde(rename = "minorFont", alias = "a:minorFont", default)]
    minor_font: Option<FontGroup>,
}

impl<'de> Deserialize<'de> for FontScheme {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = FontSchemeXml::deserialize(deserializer)?;
        Ok(FontScheme {
            major_font: xml.major_font.unwrap_or_default(),
            minor_font: xml.minor_font.unwrap_or_default(),
        })
    }
}
