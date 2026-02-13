use serde::{Deserialize, Deserializer, Serialize};

fn parse_margin_value(raw: &str) -> usize {
    raw.parse::<usize>()
        .or_else(|_| raw.parse::<f32>().map(|v| v as usize))
        .unwrap_or(0)
}

#[derive(Deserialize, Default)]
struct MarginValue {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: String,
}

fn deserialize_margin<'de, D>(deserializer: D) -> Result<usize, D::Error>
where
    D: Deserializer<'de>,
{
    let margin = Option::<MarginValue>::deserialize(deserializer)?;
    Ok(margin
        .as_ref()
        .map(|m| parse_margin_value(&m.val))
        .unwrap_or(0))
}

#[derive(Deserialize, Default)]
struct DivsChildContainer {
    #[serde(rename = "div", alias = "w:div", default)]
    div: Vec<Div>,
}

fn deserialize_divs_child<'de, D>(deserializer: D) -> Result<Vec<Div>, D::Error>
where
    D: Deserializer<'de>,
{
    let child = Option::<DivsChildContainer>::deserialize(deserializer)?;
    Ok(child.map(|c| c.div).unwrap_or_default())
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Div {
    #[serde(
        rename(serialize = "id", deserialize = "@id"),
        alias = "@w:id",
        alias = "id",
        default
    )]
    pub id: String,
    #[serde(
        rename(serialize = "marginLeft", deserialize = "marLeft"),
        alias = "w:marLeft",
        alias = "marginLeft",
        default,
        deserialize_with = "deserialize_margin"
    )]
    pub margin_left: usize,
    #[serde(
        rename(serialize = "marginRight", deserialize = "marRight"),
        alias = "w:marRight",
        alias = "marginRight",
        default,
        deserialize_with = "deserialize_margin"
    )]
    pub margin_right: usize,
    #[serde(
        rename(serialize = "marginTop", deserialize = "marTop"),
        alias = "w:marTop",
        alias = "marginTop",
        default,
        deserialize_with = "deserialize_margin"
    )]
    pub margin_top: usize,
    #[serde(
        rename(serialize = "marginBottom", deserialize = "marBottom"),
        alias = "w:marBottom",
        alias = "marginBottom",
        default,
        deserialize_with = "deserialize_margin"
    )]
    pub margin_bottom: usize,
    #[serde(
        rename(serialize = "divsChild", deserialize = "divsChild"),
        alias = "w:divsChild",
        default,
        deserialize_with = "deserialize_divs_child"
    )]
    pub divs_child: Vec<Div>,
}

impl Default for Div {
    fn default() -> Self {
        Self {
            id: "".to_string(),
            margin_left: 0,
            margin_right: 0,
            margin_top: 0,
            margin_bottom: 0,
            divs_child: vec![],
        }
    }
}

impl Div {
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            ..Default::default()
        }
    }

    pub fn margin_left(mut self, s: usize) -> Self {
        self.margin_left = s;
        self
    }

    pub fn margin_right(mut self, s: usize) -> Self {
        self.margin_right = s;
        self
    }

    pub fn margin_top(mut self, s: usize) -> Self {
        self.margin_top = s;
        self
    }

    pub fn margin_bottom(mut self, s: usize) -> Self {
        self.margin_bottom = s;
        self
    }

    pub fn add_child(mut self, s: Div) -> Self {
        self.divs_child.push(s);
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;

    #[test]
    fn test_div_json() {
        let div = Div::new("123")
            .margin_left(100)
            .margin_top(50)
            .add_child(Div::new("456").margin_right(200));
        assert_eq!(
            serde_json::to_string(&div).unwrap(),
            r#"{"id":"123","marginLeft":100,"marginRight":0,"marginTop":50,"marginBottom":0,"divsChild":[{"id":"456","marginLeft":0,"marginRight":200,"marginTop":0,"marginBottom":0,"divsChild":[]}]}"#
        );
    }
}
