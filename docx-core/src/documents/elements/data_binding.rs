use serde::{Deserialize, Serialize};
use std::io::Write;

use crate::documents::*;
use crate::xml_builder::*;

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct DataBindingXml {
    #[serde(rename = "@xpath", alias = "@w:xpath", default)]
    xpath: Option<String>,
    #[serde(rename = "@prefixMappings", alias = "@w:prefixMappings", default)]
    prefix_mappings: Option<String>,
    #[serde(rename = "@storeItemID", alias = "@w:storeItemID", default)]
    store_item_id: Option<String>,
}

#[derive(Serialize, Debug, Clone, PartialEq, Default)]
pub struct DataBinding {
    pub xpath: Option<String>,
    pub prefix_mappings: Option<String>,
    pub store_item_id: Option<String>,
}

impl<'de> Deserialize<'de> for DataBinding {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let xml = DataBindingXml::deserialize(deserializer)?;
        Ok(DataBinding {
            xpath: xml.xpath,
            prefix_mappings: xml.prefix_mappings,
            store_item_id: xml.store_item_id,
        })
    }
}

impl DataBinding {
    pub fn new() -> Self {
        Default::default()
    }

    pub fn xpath(mut self, xpath: impl Into<String>) -> Self {
        self.xpath = Some(xpath.into());
        self
    }

    pub fn prefix_mappings(mut self, m: impl Into<String>) -> Self {
        self.prefix_mappings = Some(m.into());
        self
    }

    pub fn store_item_id(mut self, id: impl Into<String>) -> Self {
        self.store_item_id = Some(id.into());
        self
    }
}

impl BuildXML for DataBinding {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .data_binding(
                self.xpath.as_ref(),
                self.prefix_mappings.as_ref(),
                self.store_item_id.as_ref(),
            )?
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
    fn test_delete_default() {
        let b = DataBinding::new().xpath("root/hello").build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:dataBinding w:xpath="root/hello" />"#
        );
    }

    #[test]
    fn test_data_binding_xml_deserialize() {
        let xml = r#"<w:dataBinding xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:xpath="root/hello" w:prefixMappings="xmlns:ns0='http://example.com'" w:storeItemID="{12345}" />"#;
        let binding: DataBinding = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(binding.xpath, Some("root/hello".to_string()));
        assert_eq!(
            binding.prefix_mappings,
            Some("xmlns:ns0='http://example.com'".to_string())
        );
        assert_eq!(binding.store_item_id, Some("{12345}".to_string()));
    }
}
