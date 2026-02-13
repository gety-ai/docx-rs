use crate::reader::ReaderError;
use std::io::Read;

pub trait FromXMLQuickXml {
    fn from_xml_quick<R: Read>(reader: R) -> Result<Self, ReaderError>
    where
        Self: std::marker::Sized;
}
