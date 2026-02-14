use quick_xml::de::from_reader;
use std::io::{BufReader, Read};

use super::*;

impl FromXMLQuickXml for CommentExtended {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Ok(from_reader(BufReader::new(reader))?)
    }
}
