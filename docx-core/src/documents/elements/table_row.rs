use serde::de::IgnoredAny;
use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;
use std::str::FromStr;

use super::{Delete, Insert, TableCell, TableRowProperty};
use crate::xml_builder::*;
use crate::{documents::BuildXML, HeightRule};

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct XmlValNode {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlWidthNode {
    #[serde(rename = "@w", alias = "@w:w", default)]
    width: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlHeightNode {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
    #[serde(rename = "@hRule", alias = "@w:hRule", default)]
    rule: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TrackChangeXml {
    #[serde(rename = "@author", alias = "@w:author", default)]
    author: Option<String>,
    #[serde(rename = "@date", alias = "@w:date", default)]
    date: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TableRowPropertyXml {
    #[serde(rename = "gridAfter", alias = "w:gridAfter", default)]
    grid_after: Option<XmlValNode>,
    #[serde(rename = "wAfter", alias = "w:wAfter", default)]
    width_after: Option<XmlWidthNode>,
    #[serde(rename = "gridBefore", alias = "w:gridBefore", default)]
    grid_before: Option<XmlValNode>,
    #[serde(rename = "wBefore", alias = "w:wBefore", default)]
    width_before: Option<XmlWidthNode>,
    #[serde(rename = "trHeight", alias = "w:trHeight", default)]
    row_height: Option<XmlHeightNode>,
    #[serde(rename = "cantSplit", alias = "w:cantSplit", default)]
    cant_split: Option<XmlValNode>,
    #[serde(rename = "ins", alias = "w:ins", default)]
    ins: Option<TrackChangeXml>,
    #[serde(rename = "del", alias = "w:del", default)]
    del: Option<TrackChangeXml>,
}

#[derive(Debug, Deserialize)]
enum TableRowChildXml {
    #[serde(rename = "tc", alias = "w:tc")]
    TableCell(TableCell),
    #[serde(rename = "trPr", alias = "w:trPr")]
    TableRowProperty(IgnoredAny),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct TableRowXml {
    #[serde(rename = "trPr", alias = "w:trPr", default)]
    property: Option<TableRowPropertyXml>,
    #[serde(rename = "$value", default)]
    children: Vec<TableRowChildXml>,
}

fn parse_on_off(v: Option<&str>) -> bool {
    !matches!(
        v.map(|x| x.trim().to_ascii_lowercase()),
        Some(ref s) if s == "0" || s == "false"
    )
}

fn parse_u32(raw: Option<String>) -> Option<u32> {
    raw.and_then(|v| v.parse::<u32>().ok())
}

fn parse_f32(raw: Option<String>) -> Option<f32> {
    raw.and_then(|v| {
        v.parse::<f32>()
            .ok()
            .or_else(|| v.parse::<f64>().ok().map(|n| n as f32))
    })
}

fn parse_insert_xml(xml: Option<TrackChangeXml>) -> Option<Insert> {
    let xml = xml?;
    let mut ins = Insert::new_with_empty();
    if let Some(author) = xml.author {
        ins = ins.author(author);
    }
    if let Some(date) = xml.date {
        ins = ins.date(date);
    }
    Some(ins)
}

fn parse_delete_xml(xml: Option<TrackChangeXml>) -> Option<Delete> {
    let xml = xml?;
    let mut del = Delete::new();
    if let Some(author) = xml.author {
        del = del.author(author);
    }
    if let Some(date) = xml.date {
        del = del.date(date);
    }
    Some(del)
}

fn parse_table_row_property_xml(xml: Option<TableRowPropertyXml>) -> TableRowProperty {
    let Some(xml) = xml else {
        return TableRowProperty::new();
    };

    let mut property = TableRowProperty::new();
    if let Some(v) = parse_u32(xml.grid_after.and_then(|v| v.val)) {
        property = property.grid_after(v);
    }
    if let Some(v) = parse_f32(xml.width_after.and_then(|v| v.width)) {
        property = property.width_after(v);
    }
    if let Some(v) = parse_u32(xml.grid_before.and_then(|v| v.val)) {
        property = property.grid_before(v);
    }
    if let Some(v) = parse_f32(xml.width_before.and_then(|v| v.width)) {
        property = property.width_before(v);
    }
    if let Some(height) = xml.row_height {
        if let Some(v) = parse_f32(height.val) {
            property = property.row_height(v);
        }
        if let Some(v) = height.rule.and_then(|v| HeightRule::from_str(&v).ok()) {
            property = property.height_rule(v);
        }
    }
    if let Some(v) = xml.cant_split {
        if parse_on_off(v.val.as_deref()) {
            property = property.cant_split();
        }
    }
    if let Some(ins) = parse_insert_xml(xml.ins) {
        property = property.insert(ins);
    }
    if let Some(del) = parse_delete_xml(xml.del) {
        property = property.delete(del);
    }
    property
}

fn table_row_child_from_xml(xml: TableRowChildXml) -> Option<TableRowChild> {
    match xml {
        TableRowChildXml::TableCell(cell) => Some(TableRowChild::TableCell(cell)),
        TableRowChildXml::TableRowProperty(_) | TableRowChildXml::Unknown => None,
    }
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    pub cells: Vec<TableRowChild>,
    pub has_numbering: bool,
    pub property: TableRowProperty,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TableRowChild {
    TableCell(TableCell),
}

impl<'de> Deserialize<'de> for TableRow {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = TableRowXml::deserialize(deserializer)?;
        let cells: Vec<TableRowChild> = xml
            .children
            .into_iter()
            .filter_map(table_row_child_from_xml)
            .collect();
        let has_numbering = cells.iter().any(|c| match c {
            TableRowChild::TableCell(cell) => cell.has_numbering,
        });

        Ok(TableRow {
            cells,
            has_numbering,
            property: parse_table_row_property_xml(xml.property),
        })
    }
}

impl BuildXML for TableRowChild {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        match self {
            TableRowChild::TableCell(v) => v.build_to(stream),
        }
    }
}

impl TableRow {
    pub fn new(cells: Vec<TableCell>) -> TableRow {
        let property = TableRowProperty::new();
        let has_numbering = cells.iter().any(|c| c.has_numbering);
        let cells = cells.into_iter().map(TableRowChild::TableCell).collect();
        Self {
            cells,
            property,
            has_numbering,
        }
    }

    pub fn grid_after(mut self, grid_after: u32) -> TableRow {
        self.property = self.property.grid_after(grid_after);
        self
    }

    pub fn width_after(mut self, w: f32) -> TableRow {
        self.property = self.property.width_after(w);
        self
    }

    pub fn grid_before(mut self, grid_before: u32) -> TableRow {
        self.property = self.property.grid_before(grid_before);
        self
    }

    pub fn width_before(mut self, w: f32) -> TableRow {
        self.property = self.property.width_before(w);
        self
    }

    pub fn row_height(mut self, h: f32) -> TableRow {
        self.property = self.property.row_height(h);
        self
    }

    pub fn height_rule(mut self, r: HeightRule) -> TableRow {
        self.property = self.property.height_rule(r);
        self
    }

    pub fn delete(mut self, d: Delete) -> TableRow {
        self.property = self.property.delete(d);
        self
    }

    pub fn insert(mut self, i: Insert) -> TableRow {
        self.property = self.property.insert(i);
        self
    }

    pub fn cant_split(mut self) -> TableRow {
        self.property = self.property.cant_split();
        self
    }
}

impl BuildXML for TableRow {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_table_row()?
            .add_child(&self.property)?
            .add_children(&self.cells)?
            .close()?
            .into_inner()
    }
}

impl Serialize for TableRowChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            TableRowChild::TableCell(ref r) => {
                let mut t = serializer.serialize_struct("TableCell", 2)?;
                t.serialize_field("type", "tableCell")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_row() {
        let b = TableRow::new(vec![TableCell::new()]).build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:tr><w:trPr /><w:tc><w:tcPr /><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr></w:p></w:tc></w:tr>"#
        );
    }

    #[test]
    fn test_row_json() {
        let r = TableRow::new(vec![TableCell::new()]);
        assert_eq!(
            serde_json::to_string(&r).unwrap(),
            r#"{"cells":[{"type":"tableCell","data":{"children":[],"property":{"width":null,"borders":null,"gridSpan":null,"verticalMerge":null,"verticalAlign":null,"textDirection":null,"shading":null},"hasNumbering":false}}],"hasNumbering":false,"property":{"gridAfter":null,"widthAfter":null,"gridBefore":null,"widthBefore":null}}"#
        );
    }

    #[test]
    fn test_row_cant_split() {
        let b = TableRow::new(vec![TableCell::new()]).cant_split().build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:tr><w:trPr><w:cantSplit /></w:trPr><w:tc><w:tcPr /><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr></w:p></w:tc></w:tr>"#
        );
    }

    #[test]
    fn test_row_xml_deserialize() {
        let xml = r#"<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:trPr>
                <w:gridAfter w:val="1"/>
                <w:wAfter w:w="100"/>
                <w:gridBefore w:val="2"/>
                <w:wBefore w:w="200"/>
                <w:trHeight w:val="500" w:hRule="exact"/>
                <w:cantSplit/>
            </w:trPr>
            <w:tc>
                <w:tcPr>
                    <w:tcW w:w="3000" w:type="dxa"/>
                </w:tcPr>
                <w:p />
            </w:tc>
        </w:tr>"#;

        let row: TableRow = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(row.cells.len(), 1);
        assert!(!row.has_numbering);

        let j = serde_json::to_value(&row).unwrap();
        assert_eq!(j["property"]["gridAfter"], 1);
        assert_eq!(j["property"]["widthAfter"], 100.0);
        assert_eq!(j["property"]["gridBefore"], 2);
        assert_eq!(j["property"]["widthBefore"], 200.0);
        assert_eq!(j["property"]["rowHeight"], 500.0);
        assert_eq!(j["property"]["heightRule"], "exact");
    }
}
