use serde::{Deserialize, Deserializer, Serialize};

use super::*;

#[derive(Debug, Clone, PartialEq, Serialize, Default)]
#[cfg_attr(feature = "wasm", derive(ts_rs::TS))]
#[cfg_attr(feature = "wasm", ts(export))]
#[serde(rename_all = "camelCase")]
pub struct Theme {
    pub font_schema: FontScheme,
}

// ============================================================================
// XML Deserialization (quick-xml serde)
// ============================================================================

#[derive(Deserialize, Default)]
struct ThemeElementsXml {
    #[serde(rename = "fontScheme", alias = "a:fontScheme", default)]
    font_scheme: Option<FontScheme>,
}

#[derive(Deserialize, Default)]
struct ThemeXml {
    #[serde(rename = "themeElements", alias = "a:themeElements", default)]
    theme_elements: ThemeElementsXml,
}

impl<'de> Deserialize<'de> for Theme {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = ThemeXml::deserialize(deserializer)?;
        Ok(Theme {
            font_schema: xml.theme_elements.font_scheme.unwrap_or_default(),
        })
    }
}
