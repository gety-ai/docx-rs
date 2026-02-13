use serde::{Deserialize, Deserializer, Serialize};

use super::*;

#[derive(Deserialize, Default)]
struct DivsContainer {
    #[serde(rename = "div", alias = "w:div", default)]
    div: Vec<Div>,
}

fn deserialize_divs<'de, D>(deserializer: D) -> Result<Vec<Div>, D::Error>
where
    D: Deserializer<'de>,
{
    let container = Option::<DivsContainer>::deserialize(deserializer)?;
    Ok(container.map(|c| c.div).unwrap_or_default())
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct WebSettings {
    #[serde(
        rename(serialize = "divs", deserialize = "divs"),
        alias = "w:divs",
        default,
        deserialize_with = "deserialize_divs"
    )]
    pub divs: Vec<Div>,
}

impl WebSettings {
    pub fn new() -> WebSettings {
        Default::default()
    }
}
