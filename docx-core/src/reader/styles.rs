use quick_xml::de::from_reader;
use std::io::{BufReader, Read};

use super::*;
use crate::reader::{FromXML, FromXMLQuickXml, ReaderError};

impl FromXMLQuickXml for Styles {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Ok(from_reader(BufReader::new(reader))?)
    }
}

impl FromXML for Styles {
    fn from_xml<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Self::from_xml_quick(reader)
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    use crate::types::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;

    #[test]
    fn test_from_xml() {
        let xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:style w:type="character" w:styleId="FootnoteTextChar">
        <w:name w:val="Footnote Text Char"></w:name>
        <w:rPr>
            <w:sz w:val="20"></w:sz>
            <w:szCs w:val="20"></w:szCs>
        </w:rPr>
        <w:uiPriority w:val="99"></w:uiPriority>
        <w:unhideWhenUsed></w:unhideWhenUsed>
        <w:basedOn w:val="DefaultParagraphFont"></w:basedOn>
        <w:link w:val="FootnoteText"></w:link>
        <w:uiPriority w:val="99"></w:uiPriority>
        <w:semiHidden></w:semiHidden>
    </w:style>
</w:styles>"#;
        let s = Styles::from_xml(xml.as_bytes()).unwrap();
        let mut styles = Styles::new();
        styles = styles.add_style(
            Style::new("FootnoteTextChar", StyleType::Character)
                .name("Footnote Text Char")
                .size(20)
                .based_on("DefaultParagraphFont")
                .link("FootnoteText"),
        );
        assert_eq!(s, styles);
    }
}
