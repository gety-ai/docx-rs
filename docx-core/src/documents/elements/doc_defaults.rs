use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;

use crate::{documents::BuildXML, RunProperty};
use crate::{xml_builder::*, LineSpacing, ParagraphProperty, ParagraphPropertyDefault};

use super::run_property_default::*;
use super::style::{
    parse_paragraph_property_xml, parse_run_property_xml, ParagraphPropertyXml, RunPropertyXml,
};
use super::RunFonts;

#[derive(Debug, Deserialize, Default)]
struct RunPropertyDefaultXml {
    #[serde(rename = "rPr", alias = "w:rPr", default)]
    run_property: Option<RunPropertyXml>,
}

#[derive(Debug, Deserialize, Default)]
struct ParagraphPropertyDefaultXml {
    #[serde(rename = "pPr", alias = "w:pPr", default)]
    paragraph_property: Option<ParagraphPropertyXml>,
}

#[derive(Debug, Deserialize, Default)]
struct DocDefaultsXml {
    #[serde(rename = "rPrDefault", alias = "w:rPrDefault", default)]
    run_property_default: Option<RunPropertyDefaultXml>,
    #[serde(rename = "pPrDefault", alias = "w:pPrDefault", default)]
    paragraph_property_default: Option<ParagraphPropertyDefaultXml>,
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocDefaults {
    run_property_default: RunPropertyDefault,
    paragraph_property_default: ParagraphPropertyDefault,
}

impl<'de> Deserialize<'de> for DocDefaults {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = DocDefaultsXml::deserialize(deserializer)?;
        let mut doc_defaults = DocDefaults::new();

        if let Some(run_defaults) = xml.run_property_default {
            let run_property = parse_run_property_xml(run_defaults.run_property);
            doc_defaults = doc_defaults.run_property(run_property);
        }
        if let Some(paragraph_defaults) = xml.paragraph_property_default {
            let paragraph_property =
                parse_paragraph_property_xml(paragraph_defaults.paragraph_property);
            doc_defaults = doc_defaults.paragraph_property(paragraph_property);
        }

        Ok(doc_defaults)
    }
}

impl DocDefaults {
    pub fn new() -> DocDefaults {
        Default::default()
    }

    pub fn size(mut self, size: usize) -> Self {
        self.run_property_default = self.run_property_default.size(size);
        self
    }

    pub fn spacing(mut self, spacing: i32) -> Self {
        self.run_property_default = self.run_property_default.spacing(spacing);
        self
    }

    pub fn fonts(mut self, font: RunFonts) -> Self {
        self.run_property_default = self.run_property_default.fonts(font);
        self
    }

    pub fn line_spacing(mut self, spacing: LineSpacing) -> Self {
        self.paragraph_property_default = self.paragraph_property_default.line_spacing(spacing);
        self
    }

    pub(crate) fn run_property(mut self, p: RunProperty) -> Self {
        self.run_property_default = self.run_property_default.run_property(p);
        self
    }

    pub(crate) fn paragraph_property(mut self, p: ParagraphProperty) -> Self {
        self.paragraph_property_default = self.paragraph_property_default.paragraph_property(p);
        self
    }
}

impl Default for DocDefaults {
    fn default() -> Self {
        let run_property_default = RunPropertyDefault::new();
        let paragraph_property_default = ParagraphPropertyDefault::new();
        DocDefaults {
            run_property_default,
            paragraph_property_default,
        }
    }
}

impl BuildXML for DocDefaults {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_doc_defaults()?
            .add_child(&self.run_property_default)?
            .add_child(&self.paragraph_property_default)?
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
    fn test_build() {
        let c = DocDefaults::new();
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:docDefaults><w:rPrDefault><w:rPr /></w:rPrDefault><w:pPrDefault><w:pPr><w:rPr /></w:pPr></w:pPrDefault></w:docDefaults>"#
        );
    }
}
