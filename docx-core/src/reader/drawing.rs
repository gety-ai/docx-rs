use quick_xml::de::from_reader;
use std::io::{BufReader, Read};

use super::*;
use crate::reader::{FromXML, FromXMLQuickXml, ReaderError};

impl FromXMLQuickXml for Drawing {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Ok(from_reader(BufReader::new(reader))?)
    }
}

impl FromXML for Drawing {
    fn from_xml<R: Read>(reader: R) -> Result<Self, ReaderError> {
        Self::from_xml_quick(reader)
    }
}
