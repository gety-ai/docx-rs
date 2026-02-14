use serde::de::IgnoredAny;
use serde::ser::{SerializeStruct, Serializer};
use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;
use std::str::FromStr;

use super::*;
use crate::documents::BuildXML;
use crate::types::*;
use crate::xml_builder::*;

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct XmlWidthNode {
    #[serde(rename = "@w", alias = "@w:w", default)]
    width: Option<String>,
    #[serde(rename = "@type", alias = "@w:type", default)]
    width_type: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlValNode {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct XmlLayoutNode {
    #[serde(rename = "@type", alias = "@w:type", default)]
    layout_type: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TablePropertyXml {
    #[serde(rename = "tblW", alias = "w:tblW", default)]
    width: Option<XmlWidthNode>,
    #[serde(rename = "jc", alias = "w:jc", default)]
    justification: Option<XmlValNode>,
    #[serde(rename = "tblInd", alias = "w:tblInd", default)]
    indent: Option<XmlWidthNode>,
    #[serde(rename = "tblStyle", alias = "w:tblStyle", default)]
    style: Option<XmlValNode>,
    #[serde(rename = "tblLayout", alias = "w:tblLayout", default)]
    layout: Option<XmlLayoutNode>,
    #[serde(rename = "tblBorders", alias = "w:tblBorders", default)]
    _borders: Option<IgnoredAny>,
    #[serde(rename = "tblCellMar", alias = "w:tblCellMar", default)]
    _margins: Option<IgnoredAny>,
}

#[derive(Debug, Deserialize, Default)]
struct GridColXml {
    #[serde(rename = "@w", alias = "@w:w", default)]
    width: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TableGridXml {
    #[serde(rename = "gridCol", alias = "w:gridCol", default)]
    columns: Vec<GridColXml>,
}

#[derive(Debug, Deserialize)]
enum TableChildXml {
    #[serde(rename = "tr", alias = "w:tr")]
    TableRow(TableRow),
    #[serde(rename = "tblPr", alias = "w:tblPr")]
    TableProperty(IgnoredAny),
    #[serde(rename = "tblGrid", alias = "w:tblGrid")]
    TableGrid(IgnoredAny),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct TableXml {
    #[serde(rename = "tblPr", alias = "w:tblPr", default)]
    property: Option<TablePropertyXml>,
    #[serde(rename = "tblGrid", alias = "w:tblGrid", default)]
    grid: Option<TableGridXml>,
    #[serde(rename = "$value", default)]
    children: Vec<TableChildXml>,
}

fn parse_usize_value(raw: Option<String>) -> Option<usize> {
    raw.and_then(|v| {
        let trimmed = v.trim().trim_end_matches('%');
        trimmed
            .parse::<usize>()
            .ok()
            .or_else(|| trimmed.parse::<f64>().ok().map(|n| n as usize))
    })
}

fn parse_table_property_xml(xml: Option<TablePropertyXml>) -> TableProperty {
    let Some(xml) = xml else {
        return TableProperty::without_borders();
    };

    let mut property = TableProperty::without_borders();
    if let Some(width) = xml.width {
        if let Some(w) = parse_usize_value(width.width) {
            let width_type = width
                .width_type
                .as_deref()
                .and_then(|s| WidthType::from_str(s).ok())
                .unwrap_or(WidthType::Auto);
            property = property.width(w, width_type);
        }
    }
    if let Some(jc) = xml.justification.and_then(|v| v.val) {
        if let Ok(v) = TableAlignmentType::from_str(&jc) {
            property = property.align(v);
        }
    }
    if let Some(ind) = xml.indent {
        if let Some(w) = parse_usize_value(ind.width) {
            property = property.indent(w as i32);
        }
    }
    if let Some(style) = xml.style.and_then(|v| v.val) {
        property = property.style(style);
    }
    if let Some(layout) = xml.layout.and_then(|v| v.layout_type) {
        if let Ok(v) = TableLayoutType::from_str(&layout) {
            property = property.layout(v);
        }
    }
    property
}

fn parse_table_grid_xml(xml: Option<TableGridXml>) -> Vec<usize> {
    let Some(xml) = xml else {
        return vec![];
    };
    xml.columns
        .into_iter()
        .filter_map(|c| parse_usize_value(c.width))
        .collect()
}

fn table_child_from_xml(xml: TableChildXml) -> Option<TableChild> {
    match xml {
        TableChildXml::TableRow(row) => Some(TableChild::TableRow(row)),
        TableChildXml::TableProperty(_) | TableChildXml::TableGrid(_) | TableChildXml::Unknown => {
            None
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Table {
    pub rows: Vec<TableChild>,
    pub grid: Vec<usize>,
    pub has_numbering: bool,
    pub property: TableProperty,
}

impl<'de> Deserialize<'de> for Table {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = TableXml::deserialize(deserializer)?;
        let rows: Vec<TableChild> = xml
            .children
            .into_iter()
            .filter_map(table_child_from_xml)
            .collect();
        let has_numbering = rows.iter().any(|r| match r {
            TableChild::TableRow(row) => row.has_numbering,
        });

        Ok(Table {
            rows,
            grid: parse_table_grid_xml(xml.grid),
            has_numbering,
            property: parse_table_property_xml(xml.property),
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TableChild {
    TableRow(TableRow),
}

impl BuildXML for TableChild {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        match self {
            TableChild::TableRow(v) => v.build_to(stream),
        }
    }
}

impl Table {
    pub fn new(rows: Vec<TableRow>) -> Table {
        let property = TableProperty::new();
        let has_numbering = rows.iter().any(|c| c.has_numbering);
        let grid = vec![];
        let rows = rows.into_iter().map(TableChild::TableRow).collect();
        Self {
            property,
            rows,
            grid,
            has_numbering,
        }
    }

    pub fn without_borders(rows: Vec<TableRow>) -> Table {
        let property = TableProperty::without_borders();
        let has_numbering = rows.iter().any(|c| c.has_numbering);
        let grid = vec![];
        let rows = rows.into_iter().map(TableChild::TableRow).collect();
        Self {
            property,
            rows,
            grid,
            has_numbering,
        }
    }

    pub fn add_row(mut self, row: TableRow) -> Table {
        self.rows.push(TableChild::TableRow(row));
        self
    }

    pub fn set_grid(mut self, grid: Vec<usize>) -> Table {
        self.grid = grid;
        self
    }

    pub fn indent(mut self, v: i32) -> Table {
        self.property = self.property.indent(v);
        self
    }

    pub fn align(mut self, v: TableAlignmentType) -> Table {
        self.property = self.property.align(v);
        self
    }

    pub fn style(mut self, s: impl Into<String>) -> Table {
        self.property = self.property.style(s);
        self
    }

    pub fn layout(mut self, t: TableLayoutType) -> Table {
        self.property = self.property.layout(t);
        self
    }

    pub fn position(mut self, p: TablePositionProperty) -> Self {
        self.property = self.property.position(p);
        self
    }

    pub fn width(mut self, w: usize, t: WidthType) -> Table {
        self.property = self.property.width(w, t);
        self
    }

    pub fn margins(mut self, margins: TableCellMargins) -> Self {
        self.property = self.property.set_margins(margins);
        self
    }

    pub fn set_borders(mut self, borders: TableBorders) -> Self {
        self.property = self.property.set_borders(borders);
        self
    }

    pub fn set_border(mut self, border: TableBorder) -> Self {
        self.property = self.property.set_border(border);
        self
    }

    pub fn clear_border(mut self, position: TableBorderPosition) -> Self {
        self.property = self.property.clear_border(position);
        self
    }

    pub fn clear_all_border(mut self) -> Self {
        self.property = self.property.clear_all_border();
        self
    }
}

impl BuildXML for Table {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        let grid = TableGrid::new(self.grid.clone());
        XMLBuilder::from(stream)
            .open_table()?
            .add_child(&self.property)?
            .add_child(&grid)?
            .add_children(&self.rows)?
            .close()?
            .into_inner()
    }
}

impl Serialize for TableChild {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            TableChild::TableRow(ref r) => {
                let mut t = serializer.serialize_struct("TableRow", 2)?;
                t.serialize_field("type", "tableRow")?;
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
    fn test_table() {
        let b = Table::new(vec![TableRow::new(vec![])]).build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto" /><w:jc w:val="left" /><w:tblBorders><w:top w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:left w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:bottom w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:right w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:insideH w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:insideV w:val="single" w:sz="2" w:space="0" w:color="000000" /></w:tblBorders></w:tblPr><w:tblGrid /><w:tr><w:trPr /></w:tr></w:tbl>"#
        );
    }

    #[test]
    fn test_table_grid() {
        let b = Table::new(vec![TableRow::new(vec![])])
            .set_grid(vec![100, 200])
            .build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto" /><w:jc w:val="left" /><w:tblBorders><w:top w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:left w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:bottom w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:right w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:insideH w:val="single" w:sz="2" w:space="0" w:color="000000" /><w:insideV w:val="single" w:sz="2" w:space="0" w:color="000000" /></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="100" w:type="dxa" /><w:gridCol w:w="200" w:type="dxa" /></w:tblGrid><w:tr><w:trPr /></w:tr></w:tbl>"#
        );
    }

    #[test]
    fn test_table_json() {
        let t = Table::new(vec![]).set_grid(vec![100, 200, 300]);
        assert_eq!(
            serde_json::to_string(&t).unwrap(),
            r#"{"rows":[],"grid":[100,200,300],"hasNumbering":false,"property":{"width":{"width":0,"widthType":"auto"},"justification":"left","borders":{"top":{"borderType":"single","size":2,"color":"000000","position":"top","space":0},"left":{"borderType":"single","size":2,"color":"000000","position":"left","space":0},"bottom":{"borderType":"single","size":2,"color":"000000","position":"bottom","space":0},"right":{"borderType":"single","size":2,"color":"000000","position":"right","space":0},"insideH":{"borderType":"single","size":2,"color":"000000","position":"insideH","space":0},"insideV":{"borderType":"single","size":2,"color":"000000","position":"insideV","space":0}}}}"#
        );
    }

    // XML Deserialization tests (quick-xml serde)
    #[test]
    fn test_table_xml_deserialize() {
        let xml = r#"<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:tblPr>
                <w:tblW w:w="9638" w:type="dxa"/>
                <w:jc w:val="center"/>
                <w:tblInd w:w="100" w:type="dxa"/>
                <w:tblStyle w:val="TableGrid"/>
                <w:tblLayout w:type="fixed"/>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="3212"/>
                <w:gridCol w:w="3213"/>
            </w:tblGrid>
            <w:tr />
            <w:tr />
        </w:tbl>"#;

        let t: Table = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(t.grid, vec![3212, 3213]);
        assert_eq!(t.rows.len(), 2);
        let j = serde_json::to_value(&t).unwrap();
        assert_eq!(j["property"]["width"]["width"], 9638);
        assert_eq!(j["property"]["width"]["widthType"], "dxa");
        assert_eq!(j["property"]["justification"], "center");
    }
}
