use super::*;
use crate::documents::BuildXML;
use crate::types::*;
use crate::xml_builder::*;
use crate::{Footer, Header};
use std::io::Write;
use std::str::FromStr;

use serde::de::IgnoredAny;
use serde::{Deserialize, Deserializer, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionProperty {
    pub page_size: PageSize,
    pub page_margin: PageMargin,
    pub columns: usize,
    pub space: usize,
    pub title_pg: bool,
    pub text_direction: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid: Option<DocGrid>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_reference: Option<HeaderReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<(String, Header)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header_reference: Option<HeaderReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_header: Option<(String, Header)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_header_reference: Option<HeaderReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_header: Option<(String, Header)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer_reference: Option<FooterReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footer: Option<(String, Footer)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_footer_reference: Option<FooterReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_footer: Option<(String, Footer)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_footer_reference: Option<FooterReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub even_footer: Option<(String, Footer)>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_type: Option<SectionType>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_num_type: Option<PageNumType>,
}

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct XmlValueAttrSP {
    #[serde(rename = "@val", alias = "@w:val", default)]
    val: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SectionPropertyXml {
    #[serde(rename = "$value", default)]
    children: Vec<SectionPropertyChildXml>,
}

#[derive(Debug, Deserialize, Default)]
struct SectionPageSizeXml {
    #[serde(rename = "@w", alias = "@w:w", default)]
    w: Option<String>,
    #[serde(rename = "@h", alias = "@w:h", default)]
    h: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SectionPageMarginXml {
    #[serde(rename = "@top", alias = "@w:top", default)]
    top: Option<String>,
    #[serde(rename = "@right", alias = "@w:right", default)]
    right: Option<String>,
    #[serde(rename = "@bottom", alias = "@w:bottom", default)]
    bottom: Option<String>,
    #[serde(rename = "@left", alias = "@w:left", default)]
    left: Option<String>,
    #[serde(rename = "@header", alias = "@w:header", default)]
    header: Option<String>,
    #[serde(rename = "@footer", alias = "@w:footer", default)]
    footer: Option<String>,
    #[serde(rename = "@gutter", alias = "@w:gutter", default)]
    gutter: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SectionDocGridXml {
    #[serde(rename = "@type", alias = "@w:type", default)]
    grid_type: Option<String>,
    #[serde(rename = "@linePitch", alias = "@w:linePitch", default)]
    line_pitch: Option<String>,
    #[serde(rename = "@charSpace", alias = "@w:charSpace", default)]
    char_space: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SectionPageNumTypeXml {
    #[serde(rename = "@start", alias = "@w:start", default)]
    start: Option<String>,
    #[serde(rename = "@chapStyle", alias = "@w:chapStyle", default)]
    chap_style: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct SectionReferenceXml {
    #[serde(rename = "@type", alias = "@w:type", default)]
    ref_type: Option<String>,
    #[serde(rename = "@id", alias = "@w:id", alias = "@r:id", default)]
    id: Option<String>,
}

#[derive(Debug, Deserialize)]
enum SectionPropertyChildXml {
    #[serde(rename = "pgMar", alias = "w:pgMar")]
    PageMargin(SectionPageMarginXml),
    #[serde(rename = "pgSz", alias = "w:pgSz")]
    PageSize(SectionPageSizeXml),
    #[serde(rename = "docGrid", alias = "w:docGrid")]
    DocGrid(SectionDocGridXml),
    #[serde(rename = "pgNumType", alias = "w:pgNumType")]
    PageNumType(SectionPageNumTypeXml),
    #[serde(rename = "headerReference", alias = "w:headerReference")]
    HeaderReference(SectionReferenceXml),
    #[serde(rename = "footerReference", alias = "w:footerReference")]
    FooterReference(SectionReferenceXml),
    #[serde(rename = "type", alias = "w:type")]
    SectionType(XmlValueAttrSP),
    #[serde(rename = "titlePg", alias = "w:titlePg")]
    TitlePg(IgnoredAny),
    #[serde(other)]
    Unknown,
}

fn parse_dxa_i32(raw: Option<String>) -> Option<i32> {
    let raw = raw?;
    let raw = raw.trim();
    if let Some(v) = raw.strip_suffix("pt") {
        v.parse::<f64>().ok().map(|n| (n * 20.0) as i32)
    } else {
        raw.parse::<f64>().ok().map(|n| n as i32)
    }
}

fn parse_dxa_u32(raw: Option<String>) -> Option<u32> {
    parse_dxa_i32(raw).and_then(|v| if v >= 0 { Some(v as u32) } else { None })
}

fn parse_doc_grid(xml: SectionDocGridXml) -> Option<DocGrid> {
    let mut doc_grid = DocGrid::with_empty();

    if let Some(grid_type) = xml.grid_type {
        if let Ok(t) = DocGridType::from_str(&grid_type) {
            doc_grid = doc_grid.grid_type(t);
        }
    }
    if let Some(line_pitch) = xml.line_pitch {
        if let Ok(lp) = line_pitch.parse::<f32>() {
            doc_grid = doc_grid.line_pitch(lp as usize);
        }
    }
    if let Some(char_space) = xml.char_space {
        if let Ok(cs) = char_space.parse::<f32>() {
            doc_grid = doc_grid.char_space(cs as isize);
        }
    }

    Some(doc_grid)
}

fn parse_page_num_type(xml: SectionPageNumTypeXml) -> PageNumType {
    let mut p = PageNumType::new();
    if let Some(start) = xml.start.and_then(|v| v.parse::<u32>().ok()) {
        p = p.start(start);
    }
    if let Some(chap_style) = xml.chap_style {
        p = p.chap_style(chap_style);
    }
    p
}

impl<'de> Deserialize<'de> for SectionProperty {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = SectionPropertyXml::deserialize(deserializer)?;
        let mut sp = SectionProperty::new();

        for child in xml.children {
            match child {
                SectionPropertyChildXml::PageMargin(v) => {
                    let mut margin = PageMargin::new();
                    if let Some(top) = parse_dxa_i32(v.top) {
                        margin = margin.top(top);
                    }
                    if let Some(right) = parse_dxa_i32(v.right) {
                        margin = margin.right(right);
                    }
                    if let Some(bottom) = parse_dxa_i32(v.bottom) {
                        margin = margin.bottom(bottom);
                    }
                    if let Some(left) = parse_dxa_i32(v.left) {
                        margin = margin.left(left);
                    }
                    if let Some(header) = parse_dxa_i32(v.header) {
                        margin = margin.header(header);
                    }
                    if let Some(footer) = parse_dxa_i32(v.footer) {
                        margin = margin.footer(footer);
                    }
                    if let Some(gutter) = parse_dxa_i32(v.gutter) {
                        margin = margin.gutter(gutter);
                    }
                    sp = sp.page_margin(margin);
                }
                SectionPropertyChildXml::PageSize(v) => {
                    let mut size = PageSize::new();
                    if let Some(w) = parse_dxa_u32(v.w) {
                        size = size.width(w);
                    }
                    if let Some(h) = parse_dxa_u32(v.h) {
                        size = size.height(h);
                    }
                    sp = sp.page_size(size);
                }
                SectionPropertyChildXml::DocGrid(v) => {
                    if let Some(doc_grid) = parse_doc_grid(v) {
                        sp = sp.doc_grid(doc_grid);
                    }
                }
                SectionPropertyChildXml::PageNumType(v) => {
                    sp = sp.page_num_type(parse_page_num_type(v));
                }
                SectionPropertyChildXml::HeaderReference(v) => {
                    let rid = v.id.unwrap_or_default();
                    let header_type = v.ref_type.unwrap_or_else(|| "default".to_string());
                    match header_type.as_str() {
                        "default" => {
                            sp.header_reference =
                                Some(HeaderReference::new(header_type, rid))
                        }
                        "first" => {
                            sp.first_header_reference =
                                Some(HeaderReference::new(header_type, rid))
                        }
                        "even" => {
                            sp.even_header_reference =
                                Some(HeaderReference::new(header_type, rid))
                        }
                        _ => {}
                    }
                }
                SectionPropertyChildXml::FooterReference(v) => {
                    let rid = v.id.unwrap_or_default();
                    let footer_type = v.ref_type.unwrap_or_else(|| "default".to_string());
                    match footer_type.as_str() {
                        "default" => {
                            sp.footer_reference =
                                Some(FooterReference::new(footer_type, rid))
                        }
                        "first" => {
                            sp.first_footer_reference =
                                Some(FooterReference::new(footer_type, rid))
                        }
                        "even" => {
                            sp.even_footer_reference =
                                Some(FooterReference::new(footer_type, rid))
                        }
                        _ => {}
                    }
                }
                SectionPropertyChildXml::SectionType(v) => {
                    if let Some(val) = v.val {
                        if let Ok(section_type) = SectionType::from_str(&val) {
                            sp.section_type = Some(section_type);
                        }
                    }
                }
                SectionPropertyChildXml::TitlePg(_) => sp = sp.title_pg(),
                SectionPropertyChildXml::Unknown => {}
            }
        }

        Ok(sp)
    }
}

impl SectionProperty {
    pub fn new() -> SectionProperty {
        Default::default()
    }

    pub fn page_size(mut self, size: PageSize) -> Self {
        self.page_size = size;
        self
    }

    pub fn page_margin(mut self, margin: PageMargin) -> Self {
        self.page_margin = margin;
        self
    }

    pub fn page_orient(mut self, o: PageOrientationType) -> Self {
        self.page_size = self.page_size.orient(o);
        self
    }

    pub fn doc_grid(mut self, doc_grid: DocGrid) -> Self {
        self.doc_grid = Some(doc_grid);
        self
    }

    pub fn text_direction(mut self, direction: String) -> Self {
        self.text_direction = direction;
        self
    }

    pub fn title_pg(mut self) -> Self {
        self.title_pg = true;
        self
    }

    pub fn header(mut self, h: Header, rid: &str) -> Self {
        self.header_reference = Some(HeaderReference::new("default", rid));
        self.header = Some((rid.to_string(), h));
        self
    }

    pub fn first_header(mut self, h: Header, rid: &str) -> Self {
        self.first_header_reference = Some(HeaderReference::new("first", rid));
        self.first_header = Some((rid.to_string(), h));
        self.title_pg = true;
        self
    }

    pub fn first_header_without_title_pg(mut self, h: Header, rid: &str) -> Self {
        self.first_header_reference = Some(HeaderReference::new("first", rid));
        self.first_header = Some((rid.to_string(), h));
        self
    }

    pub fn even_header(mut self, h: Header, rid: &str) -> Self {
        self.even_header_reference = Some(HeaderReference::new("even", rid));
        self.even_header = Some((rid.to_string(), h));
        self
    }

    pub fn footer(mut self, h: Footer, rid: &str) -> Self {
        self.footer_reference = Some(FooterReference::new("default", rid));
        self.footer = Some((rid.to_string(), h));
        self
    }

    pub fn first_footer(mut self, h: Footer, rid: &str) -> Self {
        self.first_footer_reference = Some(FooterReference::new("first", rid));
        self.first_footer = Some((rid.to_string(), h));
        self.title_pg = true;
        self
    }

    pub fn first_footer_without_title_pg(mut self, h: Footer, rid: &str) -> Self {
        self.first_footer_reference = Some(FooterReference::new("first", rid));
        self.first_footer = Some((rid.to_string(), h));
        self
    }

    pub fn even_footer(mut self, h: Footer, rid: &str) -> Self {
        self.even_footer_reference = Some(FooterReference::new("even", rid));
        self.even_footer = Some((rid.to_string(), h));
        self
    }

    pub fn get_headers(&self) -> Vec<&(String, Header)> {
        let mut headers = vec![];
        if let Some(ref header) = self.header {
            headers.push(header);
        }
        if let Some(ref header) = self.first_header {
            headers.push(header);
        }
        if let Some(ref header) = self.even_header {
            headers.push(header);
        }
        headers
    }

    pub fn get_footers(&self) -> Vec<&(String, Footer)> {
        let mut footers = vec![];
        if let Some(ref footer) = self.footer {
            footers.push(footer);
        }
        if let Some(ref footer) = self.first_footer {
            footers.push(footer);
        }
        if let Some(ref footer) = self.even_footer {
            footers.push(footer);
        }
        footers
    }

    pub fn page_num_type(mut self, h: PageNumType) -> Self {
        self.page_num_type = Some(h);
        self
    }
}

impl Default for SectionProperty {
    fn default() -> Self {
        Self {
            page_size: PageSize::new(),
            page_margin: PageMargin::new(),
            columns: 1,
            space: 425,
            title_pg: false,
            text_direction: "lrTb".to_string(),
            doc_grid: None,
            // headers
            header_reference: None,
            header: None,
            first_header_reference: None,
            first_header: None,
            even_header_reference: None,
            even_header: None,
            // footers
            footer_reference: None,
            footer: None,
            first_footer_reference: None,
            first_footer: None,
            even_footer_reference: None,
            even_footer: None,
            section_type: None,
            page_num_type: None,
        }
    }
}

impl BuildXML for SectionProperty {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        XMLBuilder::from(stream)
            .open_section_property()?
            .add_child(&self.page_size)?
            .add_child(&self.page_margin)?
            .columns(&format!("{}", &self.space), &format!("{}", &self.columns))?
            .add_optional_child(&self.doc_grid)?
            .add_optional_child(&self.header_reference)?
            .add_optional_child(&self.first_header_reference)?
            .add_optional_child(&self.even_header_reference)?
            .add_optional_child(&self.footer_reference)?
            .add_optional_child(&self.first_footer_reference)?
            .add_optional_child(&self.even_footer_reference)?
            .add_optional_child(&self.page_num_type)?
            .apply_if(self.text_direction != "lrTb", |b| {
                b.text_direction(&self.text_direction)
            })?
            .apply_opt(self.section_type, |t, b| b.type_tag(&t.to_string()))?
            .apply_if(self.title_pg, |b| b.title_pg())?
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
    fn text_section_text_direction() {
        let mut c = SectionProperty::new();
        c = c.text_direction("tbRl".to_string());
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="1" /><w:textDirection w:val="tbRl" /></w:sectPr>"#
        )
    }

    #[test]
    fn test_section_property_default() {
        let c = SectionProperty::new();
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="1" /></w:sectPr>"#
        );
    }

    #[test]
    fn test_section_property_with_footer() {
        let c = SectionProperty::new().footer(Footer::new(), "rId6");
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="1" /><w:footerReference w:type="default" r:id="rId6" /></w:sectPr>"#
        );
    }

    #[test]
    fn test_section_property_with_title_pf() {
        let c = SectionProperty::new().title_pg();
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" w:num="1" /><w:titlePg /></w:sectPr>"#
        );
    }
}
