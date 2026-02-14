use serde::Deserialize;
use std::io::{BufReader, Read};

use super::*;
use crate::reader::{FromXML, FromXMLQuickXml, ReaderError};

// ============================================================================
// XML Deserialization DTOs (quick-xml serde)
// ============================================================================

#[derive(Deserialize, Default)]
struct LpwstrXml {
    #[serde(rename = "$text", default)]
    text: String,
}

#[derive(Deserialize)]
enum PropertyChildXml {
    #[serde(rename = "lpwstr", alias = "vt:lpwstr")]
    Lpwstr(LpwstrXml),
    #[serde(other)]
    Unknown,
}

#[derive(Deserialize)]
struct PropertyXml {
    #[serde(rename = "@name", default)]
    name: String,
    #[serde(rename = "$value", default)]
    children: Vec<PropertyChildXml>,
}

#[derive(Deserialize)]
enum PropertiesChildXml {
    #[serde(rename = "property")]
    Property(PropertyXml),
    #[serde(other)]
    Unknown,
}

#[derive(Deserialize)]
struct PropertiesXml {
    #[serde(rename = "$value", default)]
    children: Vec<PropertiesChildXml>,
}

impl FromXMLQuickXml for CustomProps {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        let xml: PropertiesXml = quick_xml::de::from_reader(BufReader::new(reader))?;
        let mut props = CustomProps::new();
        for child in xml.children {
            if let PropertiesChildXml::Property(p) = child {
                if !p.name.is_empty() {
                    for pc in p.children {
                        if let PropertyChildXml::Lpwstr(v) = pc {
                            props = props.add_custom_property(&p.name, v.text);
                            break;
                        }
                    }
                }
            }
        }
        Ok(props)
    }
}

impl FromXML for CustomProps {
    fn from_xml<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Self::from_xml_quick(reader)
    }
}
