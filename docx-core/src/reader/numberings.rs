use quick_xml::de::from_reader;
use std::io::{BufReader, Read};

use super::*;
use crate::reader::{FromXML, FromXMLQuickXml, ReaderError};

impl FromXMLQuickXml for Numberings {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Ok(from_reader(BufReader::new(reader))?)
    }
}

impl FromXML for Numberings {
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
    fn test_numberings_from_xml() {
        let xml = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
    <w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">
        <w:multiLevelType w:val="hybridMultilevel"></w:multiLevelType>
        <w:lvl w:ilvl="0" w15:tentative="1">
            <w:start w:val="1"></w:start>
            <w:numFmt w:val="bullet"></w:numFmt>
            <w:lvlText w:val="●"></w:lvlText>
            <w:lvlJc w:val="left"></w:lvlJc>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"></w:ind>
            </w:pPr>
            <w:rPr></w:rPr>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"></w:abstractNumId>
    </w:num>
</w:numbering>"#;
        let n = Numberings::from_xml(xml.as_bytes()).unwrap();
        let mut nums = Numberings::new();
        let mut abs_num = AbstractNumbering::new(0).add_level(
            Level::new(
                0,
                Start::new(1),
                NumberFormat::new("bullet"),
                LevelText::new("●"),
                LevelJc::new("left"),
            )
            .indent(
                Some(720),
                Some(SpecialIndentType::Hanging(360)),
                None,
                None,
            ),
        );
        abs_num.multi_level_type = Some("hybridMultilevel".to_string());
        nums = nums
            .add_abstract_numbering(abs_num)
            .add_numbering(Numbering::new(1, 0));
        assert_eq!(n, nums)
    }

    #[test]
    fn test_numberings_from_xml_with_num_style_link() {
        let xml = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
    <w:abstractNum w:abstractNumId="0">
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:numStyleLink w:val="style1"/>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"></w:abstractNumId>
    </w:num>
</w:numbering>"#;
        let n = Numberings::from_xml(xml.as_bytes()).unwrap();
        let mut nums = Numberings::new();
        let mut abs_num = AbstractNumbering::new(0).num_style_link("style1");
        abs_num.multi_level_type = Some("hybridMultilevel".to_string());
        nums = nums
            .add_abstract_numbering(abs_num)
            .add_numbering(Numbering::new(1, 0));
        assert_eq!(n, nums)
    }

    #[test]
    fn test_numberings_from_xml_with_style_link() {
        let xml = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
    <w:abstractNum w:abstractNumId="0">
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:styleLink w:val="style1"/>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"></w:abstractNumId>
    </w:num>
</w:numbering>"#;
        let n = Numberings::from_xml(xml.as_bytes()).unwrap();
        let mut nums = Numberings::new();
        let mut abs_num = AbstractNumbering::new(0).style_link("style1");
        abs_num.multi_level_type = Some("hybridMultilevel".to_string());
        nums = nums
            .add_abstract_numbering(abs_num)
            .add_numbering(Numbering::new(1, 0));
        assert_eq!(n, nums)
    }

    #[test]
    fn test_numberings_from_xml_with_override() {
        let xml = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
    <w:abstractNum w:abstractNumId="0">
        <w:multiLevelType w:val="hybridMultilevel"/>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"></w:abstractNumId>
        <w:lvlOverride w:ilvl="0">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
        <w:lvlOverride w:ilvl="1">
          <w:startOverride w:val="1"/>
        </w:lvlOverride>
    </w:num>
</w:numbering>"#;
        let n = Numberings::from_xml(xml.as_bytes()).unwrap();
        let mut nums = Numberings::new();
        let overrides = vec![
            LevelOverride::new(0).start(1),
            LevelOverride::new(1).start(1),
        ];
        let num = Numbering::new(1, 0).overrides(overrides);
        let mut abs_num = AbstractNumbering::new(0);
        abs_num.multi_level_type = Some("hybridMultilevel".to_string());
        nums = nums.add_abstract_numbering(abs_num).add_numbering(num);
        assert_eq!(n, nums)
    }
}
