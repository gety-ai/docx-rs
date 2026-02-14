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
struct XmlBorderNode {
    #[serde(rename = "@val", alias = "@w:val", default)]
    border_type: Option<String>,
    #[serde(rename = "@sz", alias = "@w:sz", default)]
    size: Option<String>,
    #[serde(rename = "@color", alias = "@w:color", default)]
    color: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TableCellBordersXml {
    #[serde(rename = "top", alias = "w:top", default)]
    top: Option<XmlBorderNode>,
    #[serde(rename = "left", alias = "w:left", default)]
    left: Option<XmlBorderNode>,
    #[serde(rename = "bottom", alias = "w:bottom", default)]
    bottom: Option<XmlBorderNode>,
    #[serde(rename = "right", alias = "w:right", default)]
    right: Option<XmlBorderNode>,
    #[serde(rename = "insideH", alias = "w:insideH", default)]
    inside_h: Option<XmlBorderNode>,
    #[serde(rename = "insideV", alias = "w:insideV", default)]
    inside_v: Option<XmlBorderNode>,
    #[serde(rename = "tl2br", alias = "w:tl2br", default)]
    tl2br: Option<XmlBorderNode>,
    #[serde(rename = "tr2bl", alias = "w:tr2bl", default)]
    tr2bl: Option<XmlBorderNode>,
}

#[derive(Debug, Deserialize, Default)]
struct ShadingXml {
    #[serde(rename = "@val", alias = "@w:val", default)]
    shd_type: Option<String>,
    #[serde(rename = "@color", alias = "@w:color", default)]
    color: Option<String>,
    #[serde(rename = "@fill", alias = "@w:fill", default)]
    fill: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TableCellPropertyXmlHelper {
    #[serde(rename = "tcW", alias = "w:tcW", default)]
    width: Option<XmlWidthNode>,
    #[serde(rename = "gridSpan", alias = "w:gridSpan", default)]
    grid_span: Option<XmlValNode>,
    #[serde(rename = "vMerge", alias = "w:vMerge", default)]
    vertical_merge: Option<XmlValNode>,
    #[serde(rename = "vAlign", alias = "w:vAlign", default)]
    vertical_align: Option<XmlValNode>,
    #[serde(rename = "textDirection", alias = "w:textDirection", default)]
    text_direction: Option<XmlValNode>,
    #[serde(rename = "tcBorders", alias = "w:tcBorders", default)]
    borders: Option<TableCellBordersXml>,
    #[serde(rename = "shd", alias = "w:shd", default)]
    shading: Option<ShadingXml>,
}

#[derive(Debug, Deserialize)]
enum TableCellChildXml {
    #[serde(rename = "p", alias = "w:p")]
    Paragraph(Paragraph),
    #[serde(rename = "tbl", alias = "w:tbl")]
    Table(Table),
    #[serde(rename = "sdt", alias = "w:sdt")]
    StructuredDataTag(IgnoredAny),
    #[serde(rename = "tcPr", alias = "w:tcPr")]
    TableCellProperty(IgnoredAny),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct TableCellXml {
    #[serde(rename = "tcPr", alias = "w:tcPr", default)]
    property: Option<TableCellPropertyXmlHelper>,
    #[serde(rename = "$value", default)]
    children: Vec<TableCellChildXml>,
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

fn parse_table_cell_border_xml(
    node: XmlBorderNode,
    position: TableCellBorderPosition,
) -> TableCellBorder {
    let mut border = TableCellBorder::new(position);
    if let Some(v) = node
        .border_type
        .as_deref()
        .and_then(|s| BorderType::from_str(s).ok())
    {
        border = border.border_type(v);
    }
    if let Some(v) = parse_usize_value(node.size) {
        border = border.size(v);
    }
    if let Some(v) = node.color {
        border = border.color(v);
    }
    border
}

fn parse_table_cell_borders_xml(xml: Option<TableCellBordersXml>) -> Option<TableCellBorders> {
    let xml = xml?;
    let mut borders = TableCellBorders::with_empty();
    if let Some(v) = xml.top {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::Top));
    }
    if let Some(v) = xml.left {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::Left));
    }
    if let Some(v) = xml.bottom {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::Bottom));
    }
    if let Some(v) = xml.right {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::Right));
    }
    if let Some(v) = xml.inside_h {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::InsideH));
    }
    if let Some(v) = xml.inside_v {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::InsideV));
    }
    if let Some(v) = xml.tl2br {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::Tl2br));
    }
    if let Some(v) = xml.tr2bl {
        borders = borders.set(parse_table_cell_border_xml(v, TableCellBorderPosition::Tr2bl));
    }
    Some(borders)
}

fn parse_shading_xml(xml: Option<ShadingXml>) -> Option<Shading> {
    let xml = xml?;
    let mut shading = Shading::new();
    if let Some(v) = xml
        .shd_type
        .as_deref()
        .and_then(|s| ShdType::from_str(s).ok())
    {
        shading = shading.shd_type(v);
    }
    if let Some(v) = xml.color {
        shading = shading.color(v);
    }
    if let Some(v) = xml.fill {
        shading = shading.fill(v);
    }
    Some(shading)
}

fn parse_table_cell_property_xml(xml: Option<TableCellPropertyXmlHelper>) -> TableCellProperty {
    let Some(xml) = xml else {
        return TableCellProperty::new();
    };

    let mut property = TableCellProperty::new();
    if let Some(width) = xml.width {
        if let Some(v) = parse_usize_value(width.width) {
            let width_type = width
                .width_type
                .as_deref()
                .and_then(|s| WidthType::from_str(s).ok())
                .unwrap_or(WidthType::Auto);
            property = property.width(v, width_type);
        }
    }
    if let Some(v) = parse_usize_value(xml.grid_span.and_then(|v| v.val)) {
        property = property.grid_span(v);
    }
    if let Some(v) = xml.vertical_merge {
        let merge = v
            .val
            .as_deref()
            .and_then(|s| VMergeType::from_str(s).ok())
            .unwrap_or(VMergeType::Continue);
        property = property.vertical_merge(merge);
    }
    if let Some(v) = xml
        .vertical_align
        .and_then(|v| v.val)
        .and_then(|v| VAlignType::from_str(&v).ok())
    {
        property = property.vertical_align(v);
    }
    if let Some(v) = xml
        .text_direction
        .and_then(|v| v.val)
        .and_then(|v| TextDirectionType::from_str(&v).ok())
    {
        property = property.text_direction(v);
    }
    if let Some(v) = parse_table_cell_borders_xml(xml.borders) {
        property = property.set_borders(v);
    }
    if let Some(v) = parse_shading_xml(xml.shading) {
        property = property.shading(v);
    }
    property
}

fn table_cell_child_from_xml(xml: TableCellChildXml) -> Option<TableCellContent> {
    match xml {
        TableCellChildXml::Paragraph(p) => Some(TableCellContent::Paragraph(p)),
        TableCellChildXml::Table(t) => Some(TableCellContent::Table(t)),
        TableCellChildXml::StructuredDataTag(_)
        | TableCellChildXml::TableCellProperty(_)
        | TableCellChildXml::Unknown => None,
    }
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    pub children: Vec<TableCellContent>,
    pub property: TableCellProperty,
    pub has_numbering: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TableCellContent {
    Paragraph(Paragraph),
    Table(Table),
    StructuredDataTag(Box<StructuredDataTag>),
    TableOfContents(Box<TableOfContents>),
}

impl<'de> Deserialize<'de> for TableCell {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = TableCellXml::deserialize(deserializer)?;
        let children: Vec<TableCellContent> = xml
            .children
            .into_iter()
            .filter_map(table_cell_child_from_xml)
            .collect();
        let has_numbering = children.iter().any(|c| match c {
            TableCellContent::Paragraph(p) => p.has_numbering,
            TableCellContent::Table(t) => t.has_numbering,
            TableCellContent::StructuredDataTag(t) => t.has_numbering,
            TableCellContent::TableOfContents(_) => false,
        });

        Ok(TableCell {
            children,
            property: parse_table_cell_property_xml(xml.property),
            has_numbering,
        })
    }
}

impl Serialize for TableCellContent {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            TableCellContent::Paragraph(ref s) => {
                let mut t = serializer.serialize_struct("Paragraph", 2)?;
                t.serialize_field("type", "paragraph")?;
                t.serialize_field("data", s)?;
                t.end()
            }
            TableCellContent::Table(ref s) => {
                let mut t = serializer.serialize_struct("Table", 2)?;
                t.serialize_field("type", "table")?;
                t.serialize_field("data", s)?;
                t.end()
            }
            TableCellContent::StructuredDataTag(ref r) => {
                let mut t = serializer.serialize_struct("StructuredDataTag", 2)?;
                t.serialize_field("type", "structuredDataTag")?;
                t.serialize_field("data", r)?;
                t.end()
            }
            TableCellContent::TableOfContents(ref r) => {
                let mut t = serializer.serialize_struct("TableOfContents", 2)?;
                t.serialize_field("type", "tableOfContents")?;
                t.serialize_field("data", r)?;
                t.end()
            }
        }
    }
}

impl TableCell {
    pub fn new() -> TableCell {
        Default::default()
    }

    pub fn add_paragraph(mut self, p: Paragraph) -> TableCell {
        if p.has_numbering {
            self.has_numbering = true
        }
        self.children.push(TableCellContent::Paragraph(p));
        self
    }

    pub fn add_table_of_contents(mut self, t: TableOfContents) -> Self {
        self.children
            .push(TableCellContent::TableOfContents(Box::new(t)));
        self
    }

    pub fn add_structured_data_tag(mut self, t: StructuredDataTag) -> Self {
        self.children
            .push(TableCellContent::StructuredDataTag(Box::new(t)));
        self
    }

    pub fn add_table(mut self, t: Table) -> TableCell {
        if t.has_numbering {
            self.has_numbering = true
        }
        self.children.push(TableCellContent::Table(t));
        self
    }

    pub fn vertical_merge(mut self, t: VMergeType) -> TableCell {
        self.property = self.property.vertical_merge(t);
        self
    }

    pub fn shading(mut self, s: Shading) -> TableCell {
        self.property = self.property.shading(s);
        self
    }

    pub fn vertical_align(mut self, t: VAlignType) -> TableCell {
        self.property = self.property.vertical_align(t);
        self
    }

    pub fn text_direction(mut self, t: TextDirectionType) -> TableCell {
        self.property = self.property.text_direction(t);
        self
    }

    pub fn grid_span(mut self, v: usize) -> TableCell {
        self.property = self.property.grid_span(v);
        self
    }

    pub fn width(mut self, v: usize, t: WidthType) -> TableCell {
        self.property = self.property.width(v, t);
        self
    }

    pub fn set_border(mut self, border: TableCellBorder) -> Self {
        self.property = self.property.set_border(border);
        self
    }

    pub fn set_borders(mut self, borders: TableCellBorders) -> Self {
        self.property = self.property.set_borders(borders);
        self
    }

    pub fn clear_border(mut self, position: TableCellBorderPosition) -> Self {
        self.property = self.property.clear_border(position);
        self
    }

    pub fn clear_all_border(mut self) -> Self {
        self.property = self.property.clear_all_border();
        self
    }
}

impl Default for TableCell {
    fn default() -> Self {
        let property = TableCellProperty::new();
        let children = vec![];
        Self {
            property,
            children,
            has_numbering: false,
        }
    }
}

impl BuildXML for TableCell {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_table_cell()?
            .add_child(&self.property)?
            .apply_each(&self.children, |ch, b| {
                match ch {
                    TableCellContent::Paragraph(p) => b.add_child(p),
                    TableCellContent::Table(t) => {
                        b.add_child(t)?
                            // INFO: We need to add empty paragraph when parent cell includes only cell.
                            .apply_if(self.children.len() == 1, |b| b.add_child(&Paragraph::new()))
                    }
                    TableCellContent::StructuredDataTag(t) => b.add_child(&t),
                    TableCellContent::TableOfContents(t) => b.add_child(&t),
                }
            })?
            // INFO: We need to add empty paragraph when parent cell includes only cell.
            .apply_if(self.children.is_empty(), |b| b.add_child(&Paragraph::new()))?
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
    fn test_cell() {
        let b = TableCell::new().build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:tc><w:tcPr /><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr></w:p></w:tc>"#
        );
    }

    #[test]
    fn test_cell_add_p() {
        let b = TableCell::new()
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Hello")))
            .build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:tc><w:tcPr /><w:p w14:paraId="12345678"><w:pPr><w:rPr /></w:pPr><w:r><w:rPr /><w:t xml:space="preserve">Hello</w:t></w:r></w:p></w:tc>"#
        );
    }

    #[test]
    fn test_cell_json() {
        let c = TableCell::new()
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Hello")))
            .grid_span(2);
        assert_eq!(
            serde_json::to_string(&c).unwrap(),
            r#"{"children":[{"type":"paragraph","data":{"id":"12345678","children":[{"type":"run","data":{"runProperty":{},"children":[{"type":"text","data":{"preserveSpace":true,"text":"Hello"}}]}}],"property":{"runProperty":{},"tabs":[]},"hasNumbering":false}}],"property":{"width":null,"borders":null,"gridSpan":2,"verticalMerge":null,"verticalAlign":null,"textDirection":null,"shading":null},"hasNumbering":false}"#,
        );
    }

    #[test]
    fn test_cell_xml_deserialize() {
        let xml = r#"<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:tcPr>
                <w:tcW w:w="3000" w:type="dxa"/>
                <w:gridSpan w:val="2"/>
                <w:vMerge w:val="restart"/>
                <w:vAlign w:val="center"/>
                <w:tcBorders>
                    <w:top w:val="single" w:sz="8" w:color="FF0000"/>
                </w:tcBorders>
                <w:shd w:val="clear" w:fill="FFFFFF"/>
            </w:tcPr>
            <w:p />
            <w:tbl>
                <w:tr>
                    <w:tc>
                        <w:p />
                    </w:tc>
                </w:tr>
            </w:tbl>
        </w:tc>"#;

        let cell: TableCell = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(cell.children.len(), 2);
        assert!(!cell.has_numbering);

        let j = serde_json::to_value(&cell).unwrap();
        assert_eq!(j["property"]["width"]["width"], 3000);
        assert_eq!(j["property"]["width"]["widthType"], "dxa");
        assert_eq!(j["property"]["gridSpan"], 2);
        assert_eq!(j["property"]["verticalMerge"], "restart");
        assert_eq!(j["property"]["verticalAlign"], "center");
        assert_eq!(j["property"]["borders"]["top"]["borderType"], "single");
        assert_eq!(j["property"]["borders"]["top"]["size"], 8);
        assert_eq!(j["property"]["borders"]["top"]["color"], "FF0000");
        assert_eq!(j["property"]["shading"]["fill"], "FFFFFF");
    }
}
