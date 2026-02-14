use super::*;
use serde::{ser::*, Deserialize, Deserializer, Serialize};
use std::io::Write;
use std::str::FromStr;

use crate::documents::BuildXML;
use crate::types::*;
use crate::xml_builder::*;

#[derive(Debug, Clone, PartialEq, Default, Serialize)]
pub struct Drawing {
    #[serde(flatten)]
    pub data: Option<DrawingData>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum DrawingData {
    Pic(Pic),
    TextBox(TextBox),
}

impl Serialize for DrawingData {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match *self {
            DrawingData::Pic(ref pic) => {
                let mut t = serializer.serialize_struct("Pic", 2)?;
                t.serialize_field("type", "pic")?;
                t.serialize_field("data", pic)?;
                t.end()
            }
            DrawingData::TextBox(ref text_box) => {
                let mut t = serializer.serialize_struct("TextBox", 2)?;
                t.serialize_field("type", "textBox")?;
                t.serialize_field("data", text_box)?;
                t.end()
            }
        }
    }
}

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default)]
struct DrawingXml {
    #[serde(rename = "$value", default)]
    children: Vec<DrawingChildXml>,
}

#[derive(Debug, Deserialize)]
enum DrawingChildXml {
    #[serde(rename = "inline", alias = "wp:inline")]
    Inline(WpDrawingContainerXml),
    #[serde(rename = "anchor", alias = "wp:anchor")]
    Anchor(WpDrawingContainerXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct WpDrawingContainerXml {
    #[serde(rename = "@distT", alias = "@wp:distT", default)]
    dist_t: Option<String>,
    #[serde(rename = "@distB", alias = "@wp:distB", default)]
    dist_b: Option<String>,
    #[serde(rename = "@distL", alias = "@wp:distL", default)]
    dist_l: Option<String>,
    #[serde(rename = "@distR", alias = "@wp:distR", default)]
    dist_r: Option<String>,
    #[serde(rename = "@simplePos", alias = "@wp:simplePos", default)]
    simple_pos: Option<String>,
    #[serde(rename = "@layoutInCell", alias = "@wp:layoutInCell", default)]
    layout_in_cell: Option<String>,
    #[serde(rename = "@relativeHeight", alias = "@wp:relativeHeight", default)]
    relative_height: Option<String>,
    #[serde(rename = "@allowOverlap", alias = "@wp:allowOverlap", default)]
    allow_overlap: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<WpDrawingContainerChildXml>,
}

#[derive(Debug, Deserialize)]
enum WpDrawingContainerChildXml {
    #[serde(rename = "simplePos", alias = "wp:simplePos")]
    SimplePos(WpSimplePosXml),
    #[serde(rename = "positionH", alias = "wp:positionH")]
    PositionH(WpPositionXml),
    #[serde(rename = "positionV", alias = "wp:positionV")]
    PositionV(WpPositionXml),
    #[serde(rename = "extent", alias = "wp:extent")]
    Extent(WpExtentXml),
    #[serde(rename = "docPr", alias = "wp:docPr")]
    DocPr(WpDocPrXml),
    #[serde(rename = "graphic", alias = "a:graphic")]
    Graphic(AGraphicXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct DrawingXmlTextNode {
    #[serde(rename = "$text", default)]
    text: String,
}

#[derive(Debug, Deserialize, Default)]
struct WpSimplePosXml {
    #[serde(rename = "@x", alias = "@wp:x", default)]
    x: Option<String>,
    #[serde(rename = "@y", alias = "@wp:y", default)]
    y: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct WpPositionXml {
    #[serde(rename = "@relativeFrom", alias = "@wp:relativeFrom", default)]
    relative_from: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<WpPositionChildXml>,
}

#[derive(Debug, Deserialize)]
enum WpPositionChildXml {
    #[serde(rename = "posOffset", alias = "wp:posOffset")]
    PosOffset(DrawingXmlTextNode),
    #[serde(rename = "align", alias = "wp:align")]
    Align(DrawingXmlTextNode),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct WpExtentXml {
    #[serde(rename = "@cx", alias = "@wp:cx", default)]
    cx: Option<String>,
    #[serde(rename = "@cy", alias = "@wp:cy", default)]
    cy: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct WpDocPrXml {
    #[serde(rename = "@id", alias = "@wp:id", default)]
    id: Option<String>,
    #[serde(rename = "@name", alias = "@wp:name", default)]
    name: Option<String>,
    #[serde(rename = "@descr", alias = "@wp:descr", default)]
    descr: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct AGraphicXml {
    #[serde(rename = "$value", default)]
    children: Vec<AGraphicChildXml>,
}

#[derive(Debug, Deserialize)]
enum AGraphicChildXml {
    #[serde(rename = "graphicData", alias = "a:graphicData")]
    GraphicData(AGraphicDataXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct AGraphicDataXml {
    #[serde(rename = "@uri", alias = "@a:uri", default)]
    _uri: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<AGraphicDataChildXml>,
}

#[derive(Debug, Deserialize)]
enum AGraphicDataChildXml {
    #[serde(rename = "pic", alias = "pic:pic")]
    Pic(PicXml),
    #[serde(rename = "wsp", alias = "wps:wsp")]
    WpsShape(WpsShapeXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct PicXml {
    #[serde(rename = "$value", default)]
    children: Vec<PicChildXml>,
}

#[derive(Debug, Deserialize)]
enum PicChildXml {
    #[serde(rename = "nvPicPr", alias = "pic:nvPicPr")]
    NvPicPr(PicNvPicPrXml),
    #[serde(rename = "blipFill", alias = "pic:blipFill")]
    BlipFill(PicBlipFillXml),
    #[serde(rename = "spPr", alias = "pic:spPr")]
    SpPr(PicSpPrXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct PicNvPicPrXml {
    #[serde(rename = "$value", default)]
    children: Vec<PicNvPicPrChildXml>,
}

#[derive(Debug, Deserialize)]
enum PicNvPicPrChildXml {
    #[serde(rename = "cNvPr", alias = "pic:cNvPr")]
    CNvPr(PicCNvPrXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct PicCNvPrXml {
    #[serde(rename = "@name", alias = "@pic:name", default)]
    name: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct PicBlipFillXml {
    #[serde(rename = "$value", default)]
    children: Vec<PicBlipFillChildXml>,
}

#[derive(Debug, Deserialize)]
enum PicBlipFillChildXml {
    #[serde(rename = "blip", alias = "a:blip")]
    Blip(ABlipXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct ABlipXml {
    #[serde(rename = "@embed", alias = "@r:embed", default)]
    embed: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct PicSpPrXml {
    #[serde(rename = "$value", default)]
    children: Vec<PicSpPrChildXml>,
}

#[derive(Debug, Deserialize)]
enum PicSpPrChildXml {
    #[serde(rename = "xfrm", alias = "a:xfrm")]
    Xfrm(AXfrmXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct AXfrmXml {
    #[serde(rename = "@rot", alias = "@a:rot", default)]
    rot: Option<String>,
    #[serde(rename = "$value", default)]
    children: Vec<AXfrmChildXml>,
}

#[derive(Debug, Deserialize)]
enum AXfrmChildXml {
    #[serde(rename = "off", alias = "a:off")]
    Off(AOffXml),
    #[serde(rename = "ext", alias = "a:ext")]
    Ext(AExtXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct AOffXml {
    #[serde(rename = "@x", alias = "@a:x", default)]
    x: Option<String>,
    #[serde(rename = "@y", alias = "@a:y", default)]
    y: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct AExtXml {
    #[serde(rename = "@cx", alias = "@a:cx", default)]
    cx: Option<String>,
    #[serde(rename = "@cy", alias = "@a:cy", default)]
    cy: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct WpsShapeXml {
    #[serde(rename = "$value", default)]
    children: Vec<WpsShapeChildXml>,
}

#[derive(Debug, Deserialize)]
enum WpsShapeChildXml {
    #[serde(rename = "txbx", alias = "wps:txbx")]
    TextBox(WpsTextBoxXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct WpsTextBoxXml {
    #[serde(rename = "$value", default)]
    children: Vec<WpsTextBoxChildXml>,
}

#[derive(Debug, Deserialize)]
enum WpsTextBoxChildXml {
    #[serde(rename = "txbxContent", alias = "w:txbxContent")]
    Content(TextBoxContentXml),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Deserialize, Default)]
struct TextBoxContentXml {
    #[serde(rename = "$value", default)]
    children: Vec<TextBoxContentChildXml>,
}

#[derive(Debug, Deserialize)]
enum TextBoxContentChildXml {
    #[serde(rename = "p", alias = "w:p")]
    Paragraph(Paragraph),
    #[serde(rename = "tbl", alias = "w:tbl")]
    Table(Table),
    #[serde(other)]
    Unknown,
}

fn parse_i32_value(raw: Option<String>) -> Option<i32> {
    raw.and_then(|v| {
        let trimmed = v.trim();
        trimmed
            .parse::<i32>()
            .ok()
            .or_else(|| trimmed.parse::<f64>().ok().map(|n| n as i32))
    })
}

fn parse_u32_value(raw: Option<String>) -> Option<u32> {
    raw.and_then(|v| {
        let trimmed = v.trim();
        trimmed
            .parse::<u32>()
            .ok()
            .or_else(|| trimmed.parse::<f64>().ok().map(|n| n as u32))
    })
}

fn parse_on_off(raw: Option<String>, default: bool) -> bool {
    match raw.as_deref().map(|v| v.trim().to_ascii_lowercase()) {
        Some(v) if matches!(v.as_str(), "0" | "false" | "off") => false,
        Some(v) if matches!(v.as_str(), "1" | "true" | "on") => true,
        Some(_) => true,
        None => default,
    }
}

fn parse_pic_align(raw: &str) -> Option<PicAlign> {
    match raw.trim().to_ascii_lowercase().as_str() {
        "left" => Some(PicAlign::Left),
        "right" => Some(PicAlign::Right),
        "center" => Some(PicAlign::Center),
        "bottom" => Some(PicAlign::Bottom),
        "top" => Some(PicAlign::Top),
        _ => None,
    }
}

fn parse_wp_position(children: Vec<WpPositionChildXml>) -> DrawingPosition {
    for child in children {
        match child {
            WpPositionChildXml::PosOffset(node) => {
                if let Some(offset) = parse_i32_value(Some(node.text)) {
                    return DrawingPosition::Offset(offset);
                }
            }
            WpPositionChildXml::Align(node) => {
                if let Some(align) = parse_pic_align(&node.text) {
                    return DrawingPosition::Align(align);
                }
            }
            WpPositionChildXml::Unknown => {}
        }
    }
    DrawingPosition::Offset(0)
}

fn parse_text_box_content_children(xml: TextBoxContentXml) -> Vec<TextBoxContentChild> {
    xml.children
        .into_iter()
        .filter_map(|child| match child {
            TextBoxContentChildXml::Paragraph(p) => {
                Some(TextBoxContentChild::Paragraph(Box::new(p)))
            }
            TextBoxContentChildXml::Table(t) => Some(TextBoxContentChild::Table(Box::new(t))),
            TextBoxContentChildXml::Unknown => None,
        })
        .collect()
}

fn parse_wps_shape_text_box(xml: WpsShapeXml) -> Option<Vec<TextBoxContentChild>> {
    for child in xml.children {
        if let WpsShapeChildXml::TextBox(tbx) = child {
            for tbx_child in tbx.children {
                if let WpsTextBoxChildXml::Content(content) = tbx_child {
                    return Some(parse_text_box_content_children(content));
                }
            }
        }
    }
    None
}

fn parse_pic_from_xml(xml: PicXml) -> Pic {
    let mut pic = Pic::with_empty();

    for child in xml.children {
        match child {
            PicChildXml::NvPicPr(nv_pic_pr) => {
                for nv_child in nv_pic_pr.children {
                    if let PicNvPicPrChildXml::CNvPr(c_nv_pr) = nv_child {
                        if let Some(name) = c_nv_pr.name {
                            if pic.name.is_empty() {
                                pic.name = name;
                            }
                        }
                    }
                }
            }
            PicChildXml::BlipFill(blip_fill) => {
                for blip_child in blip_fill.children {
                    if let PicBlipFillChildXml::Blip(blip) = blip_child {
                        if let Some(embed) = blip.embed {
                            pic.id = embed;
                        }
                    }
                }
            }
            PicChildXml::SpPr(sp_pr) => {
                for sp_child in sp_pr.children {
                    if let PicSpPrChildXml::Xfrm(xfrm) = sp_child {
                        if let Some(rot) = parse_u32_value(xfrm.rot) {
                            pic.rot = (rot / 60_000) as u16;
                        }
                        for xfrm_child in xfrm.children {
                            match xfrm_child {
                                AXfrmChildXml::Off(off) => {
                                    pic.position_h =
                                        DrawingPosition::Offset(parse_i32_value(off.x).unwrap_or(0));
                                    pic.position_v =
                                        DrawingPosition::Offset(parse_i32_value(off.y).unwrap_or(0));
                                }
                                AXfrmChildXml::Ext(ext) => {
                                    pic.size = (
                                        parse_u32_value(ext.cx).unwrap_or(0),
                                        parse_u32_value(ext.cy).unwrap_or(0),
                                    );
                                }
                                AXfrmChildXml::Unknown => {}
                            }
                        }
                    }
                }
            }
            PicChildXml::Unknown => {}
        }
    }

    pic
}

fn parse_graphic_payload(xml: AGraphicXml) -> (Option<Pic>, Option<Vec<TextBoxContentChild>>) {
    let mut pic = None;
    let mut text_box = None;

    for child in xml.children {
        if let AGraphicChildXml::GraphicData(data) = child {
            for data_child in data.children {
                match data_child {
                    AGraphicDataChildXml::Pic(pic_xml) if pic.is_none() => {
                        pic = Some(parse_pic_from_xml(pic_xml));
                    }
                    AGraphicDataChildXml::WpsShape(shape_xml) if text_box.is_none() => {
                        text_box = parse_wps_shape_text_box(shape_xml);
                    }
                    _ => {}
                }
            }
        }
    }

    (pic, text_box)
}

fn parse_drawing_container(
    xml: WpDrawingContainerXml,
    position_type: DrawingPositionType,
) -> Drawing {
    let simple_pos = parse_on_off(xml.simple_pos, false);
    let mut simple_pos_x = 0;
    let mut simple_pos_y = 0;
    let layout_in_cell = parse_on_off(xml.layout_in_cell, true);
    let relative_height = parse_u32_value(xml.relative_height).unwrap_or(0);
    let allow_overlap = parse_on_off(xml.allow_overlap, true);
    let dist_t = parse_i32_value(xml.dist_t).unwrap_or(0);
    let dist_b = parse_i32_value(xml.dist_b).unwrap_or(0);
    let dist_l = parse_i32_value(xml.dist_l).unwrap_or(0);
    let dist_r = parse_i32_value(xml.dist_r).unwrap_or(0);

    let mut relative_from_h = RelativeFromHType::default();
    let mut relative_from_v = RelativeFromVType::default();
    let mut position_h = DrawingPosition::Offset(0);
    let mut position_v = DrawingPosition::Offset(0);
    let mut extent = (0_u32, 0_u32);
    let mut doc_pr_id = String::new();
    let mut doc_pr_name = String::new();
    let mut doc_pr_descr = String::new();

    let mut pic = None;
    let mut text_box_children = None;

    for child in xml.children {
        match child {
            WpDrawingContainerChildXml::SimplePos(node) => {
                simple_pos_x = parse_i32_value(node.x).unwrap_or(0);
                simple_pos_y = parse_i32_value(node.y).unwrap_or(0);
            }
            WpDrawingContainerChildXml::PositionH(node) => {
                if let Some(v) = node
                    .relative_from
                    .as_deref()
                    .and_then(|v| RelativeFromHType::from_str(v).ok())
                {
                    relative_from_h = v;
                }
                position_h = parse_wp_position(node.children);
            }
            WpDrawingContainerChildXml::PositionV(node) => {
                if let Some(v) = node
                    .relative_from
                    .as_deref()
                    .and_then(|v| RelativeFromVType::from_str(v).ok())
                {
                    relative_from_v = v;
                }
                position_v = parse_wp_position(node.children);
            }
            WpDrawingContainerChildXml::Extent(node) => {
                extent = (
                    parse_u32_value(node.cx).unwrap_or(0),
                    parse_u32_value(node.cy).unwrap_or(0),
                );
            }
            WpDrawingContainerChildXml::DocPr(node) => {
                if let Some(v) = node.id {
                    doc_pr_id = v;
                }
                if let Some(v) = node.name {
                    doc_pr_name = v;
                }
                if let Some(v) = node.descr {
                    doc_pr_descr = v;
                }
            }
            WpDrawingContainerChildXml::Graphic(node) => {
                let (parsed_pic, parsed_text_box) = parse_graphic_payload(node);
                if pic.is_none() {
                    pic = parsed_pic;
                }
                if text_box_children.is_none() {
                    text_box_children = parsed_text_box;
                }
            }
            WpDrawingContainerChildXml::Unknown => {}
        }
    }

    if let Some(mut pic) = pic {
        pic.position_type = position_type;
        pic.simple_pos = simple_pos;
        pic.simple_pos_x = simple_pos_x;
        pic.simple_pos_y = simple_pos_y;
        pic.layout_in_cell = layout_in_cell;
        pic.relative_height = relative_height;
        pic.allow_overlap = allow_overlap;
        pic.dist_t = dist_t;
        pic.dist_b = dist_b;
        pic.dist_l = dist_l;
        pic.dist_r = dist_r;
        pic.relative_from_h = relative_from_h;
        pic.relative_from_v = relative_from_v;
        pic.position_h = position_h;
        pic.position_v = position_v;
        pic.doc_pr_id = doc_pr_id;
        pic.name = doc_pr_name;
        pic.description = doc_pr_descr;
        if pic.size == (0, 0) && extent != (0, 0) {
            pic.size = extent;
        }
        return Drawing::new().pic(pic);
    }

    if let Some(children) = text_box_children {
        let mut text_box = TextBox::new();
        text_box.position_type = position_type;
        text_box.simple_pos = simple_pos;
        text_box.simple_pos_x = simple_pos_x;
        text_box.simple_pos_y = simple_pos_y;
        text_box.layout_in_cell = layout_in_cell;
        text_box.relative_height = relative_height;
        text_box.allow_overlap = allow_overlap;
        text_box.dist_t = dist_t;
        text_box.dist_b = dist_b;
        text_box.dist_l = dist_l;
        text_box.dist_r = dist_r;
        text_box.relative_from_h = relative_from_h;
        text_box.relative_from_v = relative_from_v;
        text_box.position_h = position_h;
        text_box.position_v = position_v;
        text_box.children = children;
        if extent != (0, 0) {
            text_box.size = extent;
        }
        return Drawing::new().text_box(text_box);
    }

    Drawing::new()
}

impl<'de> Deserialize<'de> for Drawing {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = DrawingXml::deserialize(deserializer)?;

        for child in xml.children {
            match child {
                DrawingChildXml::Inline(inline) => {
                    return Ok(parse_drawing_container(inline, DrawingPositionType::Inline))
                }
                DrawingChildXml::Anchor(anchor) => {
                    return Ok(parse_drawing_container(anchor, DrawingPositionType::Anchor))
                }
                DrawingChildXml::Unknown => {}
            }
        }

        Ok(Drawing::new())
    }
}

impl Drawing {
    pub fn new() -> Drawing {
        Default::default()
    }

    pub fn pic(mut self, pic: Pic) -> Drawing {
        self.data = Some(DrawingData::Pic(pic));
        self
    }

    pub fn text_box(mut self, t: TextBox) -> Drawing {
        self.data = Some(DrawingData::TextBox(t));
        self
    }
}

impl BuildXML for Drawing {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        let b = XMLBuilder::from(stream);
        let mut b = b.open_drawing()?;

        match &self.data {
            Some(DrawingData::Pic(p)) => {
                if let DrawingPositionType::Inline { .. } = p.position_type {
                    b = b.open_wp_inline(
                        &format!("{}", p.dist_t),
                        &format!("{}", p.dist_b),
                        &format!("{}", p.dist_l),
                        &format!("{}", p.dist_r),
                    )?
                } else {
                    b = b
                        .open_wp_anchor(
                            &format!("{}", p.dist_t),
                            &format!("{}", p.dist_b),
                            &format!("{}", p.dist_l),
                            &format!("{}", p.dist_r),
                            "0",
                            if p.simple_pos { "1" } else { "0" },
                            "0",
                            "0",
                            if p.layout_in_cell { "1" } else { "0" },
                            &format!("{}", p.relative_height),
                        )?
                        .simple_pos(
                            &format!("{}", p.simple_pos_x),
                            &format!("{}", p.simple_pos_y),
                        )?
                        .open_position_h(&format!("{}", p.relative_from_h))?;

                    match p.position_h {
                        DrawingPosition::Offset(x) => {
                            let x = format!("{}", x as u32);
                            b = b.pos_offset(&x)?.close()?;
                        }
                        DrawingPosition::Align(x) => {
                            b = b.align(&x.to_string())?.close()?;
                        }
                    }

                    b = b.open_position_v(&format!("{}", p.relative_from_v))?;

                    match p.position_v {
                        DrawingPosition::Offset(y) => {
                            let y = format!("{}", y as u32);
                            b = b.pos_offset(&y)?.close()?;
                        }
                        DrawingPosition::Align(a) => {
                            b = b.align(&a.to_string())?.close()?;
                        }
                    }
                }

                let w = format!("{}", p.size.0);
                let h = format!("{}", p.size.1);
                b = b
                    // Please see 20.4.2.7 extent (Drawing Object Size)
                    // One inch equates to 914400 EMUs and a centimeter is 360000
                    .wp_extent(&w, &h)?
                    .wp_effect_extent("0", "0", "0", "0")?;
                if p.allow_overlap {
                    b = b.wrap_none()?;
                } else if p.position_type == DrawingPositionType::Anchor {
                    b = b.wrap_square("bothSides")?;
                }
                let doc_pr_id_str = if p.doc_pr_id.is_empty() { "1" } else { &p.doc_pr_id };
                b = b
                    .wp_doc_pr(doc_pr_id_str, p.name_or_default(), &p.description)?
                    .open_wp_c_nv_graphic_frame_pr()?
                    .a_graphic_frame_locks(
                        "http://schemas.openxmlformats.org/drawingml/2006/main",
                        "1",
                    )?
                    .close()?
                    .open_a_graphic("http://schemas.openxmlformats.org/drawingml/2006/main")?
                    .open_a_graphic_data(
                        "http://schemas.openxmlformats.org/drawingml/2006/picture",
                    )?
                    .add_child(&p.clone())?
                    .close()?
                    .close()?;
            }
            Some(DrawingData::TextBox(_t)) => unimplemented!("TODO: Support textBox writer"),
            None => {
                unimplemented!()
            }
        }
        b.close()?.close()?.into_inner()
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_drawing_build_with_pic() {
        let pic = Pic::new_with_dimensions(Vec::new(), 320, 240);
        let d = Drawing::new().pic(pic).build();
        assert_eq!(
            str::from_utf8(&d).unwrap(),
            r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3048000" cy="2286000" /><wp:effectExtent b="0" l="0" r="0" t="0" /><wp:docPr id="1" name="Figure" descr="" /><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" /></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="" /><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1" /></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage123" /><a:srcRect /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm rot="0"><a:off x="0" y="0" /><a:ext cx="3048000" cy="2286000" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#
        );
    }

    #[test]
    fn test_drawing_build_with_pic_overlap() {
        let pic = Pic::new_with_dimensions(Vec::new(), 320, 240).overlapping();
        let d = Drawing::new().pic(pic).build();
        assert_eq!(
            str::from_utf8(&d).unwrap(),
            r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3048000" cy="2286000" /><wp:effectExtent b="0" l="0" r="0" t="0" /><wp:wrapNone /><wp:docPr id="1" name="Figure" descr="" /><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" /></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="" /><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1" /></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage123" /><a:srcRect /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm rot="0"><a:off x="0" y="0" /><a:ext cx="3048000" cy="2286000" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#
        );
    }

    #[test]
    fn test_drawing_build_with_pic_align_right() {
        let mut pic = Pic::new_with_dimensions(Vec::new(), 320, 240).floating();
        pic = pic.relative_from_h(RelativeFromHType::Column);
        pic = pic.relative_from_v(RelativeFromVType::Paragraph);
        pic = pic.position_h(DrawingPosition::Align(PicAlign::Right));
        let d = Drawing::new().pic(pic).build();
        assert_eq!(
            str::from_utf8(&d).unwrap(),
            r#"<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="0" behindDoc="0" locked="0" layoutInCell="0" relativeHeight="190500"><wp:simplePos x="0" y="0" /><wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="3048000" cy="2286000" /><wp:effectExtent b="0" l="0" r="0" t="0" /><wp:wrapSquare wrapText="bothSides" /><wp:docPr id="1" name="Figure" descr="" /><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" /></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="" /><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1" /></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage123" /><a:srcRect /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm rot="0"><a:off x="0" y="0" /><a:ext cx="3048000" cy="2286000" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"#
        );
    }

    #[test]
    fn test_issue686() {
        let pic = Pic::new_with_dimensions(Vec::new(), 320, 240)
            .size(320 * 9525, 240 * 9525)
            .floating()
            .offset_x(300 * 9525)
            .offset_y(400 * 9525);

        let d = Drawing::new().pic(pic).build();
        assert_eq!(
            str::from_utf8(&d).unwrap(),
            r#"<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="0" behindDoc="0" locked="0" layoutInCell="0" relativeHeight="190500"><wp:simplePos x="0" y="0" /><wp:positionH relativeFrom="margin"><wp:posOffset>2857500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:posOffset>3810000</wp:posOffset></wp:positionV><wp:extent cx="3048000" cy="2286000" /><wp:effectExtent b="0" l="0" r="0" t="0" /><wp:wrapSquare wrapText="bothSides" /><wp:docPr id="1" name="Figure" descr="" /><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" /></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="" /><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1" /></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage123" /><a:srcRect /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm rot="0"><a:off x="0" y="0" /><a:ext cx="3048000" cy="2286000" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"#
        );
    }
}
