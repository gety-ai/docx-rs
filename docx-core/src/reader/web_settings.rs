use quick_xml::de::from_reader;
use std::io::{BufReader, Read};

use super::*;
use crate::reader::{FromXML, FromXMLQuickXml, ReaderError};

impl FromXMLQuickXml for WebSettings {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Ok(from_reader(BufReader::new(reader))?)
    }
}

impl FromXML for WebSettings {
    fn from_xml<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Self::from_xml_quick(reader)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;

    #[test]
    fn test_read_web_settings_xml() {
        let xml = r#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:divs>
                <w:div w:id="1234">
                    <w:marLeft w:val="100"/>
                    <w:marRight w:val="200"/>
                    <w:marTop w:val="50"/>
                    <w:marBottom w:val="75"/>
                    <w:divsChild>
                        <w:div w:id="5678">
                            <w:marTop w:val="25"/>
                        </w:div>
                    </w:divsChild>
                </w:div>
            </w:divs>
        </w:webSettings>"#;

        let ws = WebSettings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(ws.divs.len(), 1);
        let div = &ws.divs[0];
        assert_eq!(div.id, "1234");
        assert_eq!(div.margin_left, 100);
        assert_eq!(div.margin_right, 200);
        assert_eq!(div.margin_top, 50);
        assert_eq!(div.margin_bottom, 75);
        assert_eq!(div.divs_child.len(), 1);
        assert_eq!(div.divs_child[0].id, "5678");
        assert_eq!(div.divs_child[0].margin_top, 25);
    }

    #[test]
    fn test_read_web_settings_without_id() {
        let xml = r#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:divs>
                <w:div>
                    <w:marLeft w:val="100"/>
                </w:div>
            </w:divs>
        </w:webSettings>"#;

        let ws = WebSettings::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(ws.divs.len(), 1);
        assert_eq!(ws.divs[0].id, "");
        assert_eq!(ws.divs[0].margin_left, 100);
    }
}
