use crate::documents::BuildXML;
use crate::xml_builder::*;
use std::io::Write;

use serde::{Deserialize, Deserializer, Serialize, Serializer};

#[derive(Debug, Clone, PartialEq)]
pub struct Zoom {
    val: usize,
}

// XML deserialization helper
#[derive(Deserialize)]
struct ZoomXml {
    #[serde(rename = "@val", alias = "@w:val", alias = "@percent", alias = "@w:percent", default)]
    val: Option<String>,
}

impl<'de> Deserialize<'de> for Zoom {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = ZoomXml::deserialize(deserializer)?;
        let val = xml
            .val
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(100);
        Ok(Zoom { val })
    }
}

impl Zoom {
    pub fn new(val: usize) -> Zoom {
        Zoom { val }
    }
}

impl BuildXML for Zoom {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .zoom(&format!("{}", self.val))?
            .into_inner()
    }
}

impl Serialize for Zoom {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_u64(self.val as u64)
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_zoom() {
        let c = Zoom::new(20);
        let b = c.build();
        assert_eq!(str::from_utf8(&b).unwrap(), r#"<w:zoom w:percent="20" />"#);
    }
}
