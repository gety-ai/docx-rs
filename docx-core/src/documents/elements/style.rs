use serde::{Deserialize, Deserializer, Serialize};
use std::io::Write;
use std::str::FromStr;

use crate::documents::BuildXML;
use crate::escape::escape;
use crate::types::*;
use crate::xml_builder::*;
use crate::StyleType;

use super::*;

// ============================================================================
// XML Deserialization Helper Structures (for quick-xml serde)
// ============================================================================

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct XmlValueAttr {
    #[serde(rename = "@val", alias = "@w:val", default)]
    pub val: Option<String>,
}

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct XmlOnOffAttr {
    #[serde(rename = "@val", alias = "@w:val", default)]
    pub val: Option<String>,
}

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct RunFontsXml {
    #[serde(rename = "@ascii", alias = "@w:ascii", default)]
    pub ascii: Option<String>,
    #[serde(rename = "@eastAsia", alias = "@w:eastAsia", default)]
    pub east_asia: Option<String>,
    #[serde(rename = "@hAnsi", alias = "@w:hAnsi", default)]
    pub h_ansi: Option<String>,
    #[serde(rename = "@cs", alias = "@w:cs", default)]
    pub cs: Option<String>,
    #[serde(rename = "@asciiTheme", alias = "@w:asciiTheme", default)]
    pub ascii_theme: Option<String>,
    #[serde(rename = "@eastAsiaTheme", alias = "@w:eastAsiaTheme", default)]
    pub east_asia_theme: Option<String>,
    #[serde(rename = "@hAnsiTheme", alias = "@w:hAnsiTheme", default)]
    pub h_ansi_theme: Option<String>,
    #[serde(rename = "@cstheme", alias = "@w:cstheme", default)]
    pub cs_theme: Option<String>,
    #[serde(rename = "@hint", alias = "@w:hint", default)]
    pub hint: Option<String>,
}

// Internal wrapper for deserializing rPr with potentially duplicate elements
#[derive(Debug, Deserialize, Default)]
struct RunPropertyXmlRaw {
    #[serde(rename = "$value", default)]
    children: Vec<RunPropertyChildXml>,
}

#[derive(Debug, Deserialize)]
enum RunPropertyChildXml {
    #[serde(rename = "rStyle", alias = "w:rStyle")]
    Style(XmlValueAttr),
    #[serde(rename = "sz", alias = "w:sz")]
    Size(XmlValueAttr),
    #[serde(rename = "color", alias = "w:color")]
    Color(XmlValueAttr),
    #[serde(rename = "highlight", alias = "w:highlight")]
    Highlight(XmlValueAttr),
    #[serde(rename = "spacing", alias = "w:spacing")]
    Spacing(XmlValueAttr),
    #[serde(rename = "rFonts", alias = "w:rFonts")]
    Fonts(RunFontsXml),
    #[serde(rename = "u", alias = "w:u")]
    Underline(XmlValueAttr),
    #[serde(rename = "b", alias = "w:b")]
    Bold(XmlOnOffAttr),
    #[serde(rename = "bCs", alias = "w:bCs")]
    BoldCs(XmlOnOffAttr),
    #[serde(rename = "i", alias = "w:i")]
    Italic(XmlOnOffAttr),
    #[serde(rename = "iCs", alias = "w:iCs")]
    ItalicCs(XmlOnOffAttr),
    #[serde(rename = "strike", alias = "w:strike")]
    Strike(XmlOnOffAttr),
    #[serde(rename = "dstrike", alias = "w:dstrike")]
    Dstrike(XmlOnOffAttr),
    #[serde(rename = "vanish", alias = "w:vanish")]
    Vanish(XmlOnOffAttr),
    #[serde(rename = "specVanish", alias = "w:specVanish")]
    SpecVanish(XmlOnOffAttr),
    #[serde(other)]
    Unknown,
}

#[derive(Debug, Default, Clone)]
pub(crate) struct RunPropertyXml {
    pub style: Option<XmlValueAttr>,
    pub size: Option<XmlValueAttr>,
    pub color: Option<XmlValueAttr>,
    pub highlight: Option<XmlValueAttr>,
    pub spacing: Option<XmlValueAttr>,
    pub fonts: Option<RunFontsXml>,
    pub underline: Option<XmlValueAttr>,
    pub bold: Option<XmlOnOffAttr>,
    pub bold_cs: Option<XmlOnOffAttr>,
    pub italic: Option<XmlOnOffAttr>,
    pub italic_cs: Option<XmlOnOffAttr>,
    pub strike: Option<XmlOnOffAttr>,
    pub dstrike: Option<XmlOnOffAttr>,
    pub vanish: Option<XmlOnOffAttr>,
    pub spec_vanish: Option<XmlOnOffAttr>,
}

impl<'de> Deserialize<'de> for RunPropertyXml {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let raw = RunPropertyXmlRaw::deserialize(deserializer)?;
        let mut result = RunPropertyXml::default();

        for child in raw.children {
            match child {
                RunPropertyChildXml::Style(v) if result.style.is_none() => result.style = Some(v),
                RunPropertyChildXml::Size(v) if result.size.is_none() => result.size = Some(v),
                RunPropertyChildXml::Color(v) if result.color.is_none() => result.color = Some(v),
                RunPropertyChildXml::Highlight(v) if result.highlight.is_none() => {
                    result.highlight = Some(v)
                }
                RunPropertyChildXml::Spacing(v) if result.spacing.is_none() => {
                    result.spacing = Some(v)
                }
                RunPropertyChildXml::Fonts(v) if result.fonts.is_none() => result.fonts = Some(v),
                RunPropertyChildXml::Underline(v) if result.underline.is_none() => {
                    result.underline = Some(v)
                }
                RunPropertyChildXml::Bold(v) if result.bold.is_none() => result.bold = Some(v),
                RunPropertyChildXml::BoldCs(v) if result.bold_cs.is_none() => {
                    result.bold_cs = Some(v)
                }
                RunPropertyChildXml::Italic(v) if result.italic.is_none() => result.italic = Some(v),
                RunPropertyChildXml::ItalicCs(v) if result.italic_cs.is_none() => {
                    result.italic_cs = Some(v)
                }
                RunPropertyChildXml::Strike(v) if result.strike.is_none() => result.strike = Some(v),
                RunPropertyChildXml::Dstrike(v) if result.dstrike.is_none() => {
                    result.dstrike = Some(v)
                }
                RunPropertyChildXml::Vanish(v) if result.vanish.is_none() => result.vanish = Some(v),
                RunPropertyChildXml::SpecVanish(v) if result.spec_vanish.is_none() => {
                    result.spec_vanish = Some(v)
                }
                _ => {} // Ignore duplicates and unknown elements
            }
        }

        Ok(result)
    }
}

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct IndentXml {
    #[serde(
        rename = "@left",
        alias = "@w:left",
        alias = "@start",
        alias = "@w:start",
        default
    )]
    pub left: Option<String>,
    #[serde(
        rename = "@right",
        alias = "@w:right",
        alias = "@end",
        alias = "@w:end",
        default
    )]
    pub right: Option<String>,
    #[serde(rename = "@hanging", alias = "@w:hanging", default)]
    pub hanging: Option<String>,
    #[serde(rename = "@firstLine", alias = "@w:firstLine", default)]
    pub first_line: Option<String>,
    #[serde(rename = "@startChars", alias = "@w:startChars", default)]
    pub start_chars: Option<String>,
    #[serde(rename = "@hangingChars", alias = "@w:hangingChars", default)]
    pub hanging_chars: Option<String>,
    #[serde(rename = "@firstLineChars", alias = "@w:firstLineChars", default)]
    pub first_line_chars: Option<String>,
}

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct LineSpacingXml {
    #[serde(rename = "@lineRule", alias = "@w:lineRule", default)]
    pub line_rule: Option<String>,
    #[serde(rename = "@before", alias = "@w:before", default)]
    pub before: Option<String>,
    #[serde(rename = "@after", alias = "@w:after", default)]
    pub after: Option<String>,
    #[serde(rename = "@beforeLines", alias = "@w:beforeLines", default)]
    pub before_lines: Option<String>,
    #[serde(rename = "@afterLines", alias = "@w:afterLines", default)]
    pub after_lines: Option<String>,
    #[serde(rename = "@line", alias = "@w:line", default)]
    pub line: Option<String>,
}

#[derive(Debug, Deserialize, Default, Clone)]
pub(crate) struct ParagraphPropertyXml {
    #[serde(rename = "rPr", alias = "w:rPr", default)]
    pub run_property: Option<RunPropertyXml>,
    #[serde(rename = "pStyle", alias = "w:pStyle", default)]
    pub style: Option<XmlValueAttr>,
    #[serde(rename = "jc", alias = "w:jc", default)]
    pub alignment: Option<XmlValueAttr>,
    #[serde(rename = "ind", alias = "w:ind", default)]
    pub indent: Option<IndentXml>,
    #[serde(rename = "spacing", alias = "w:spacing", default)]
    pub spacing: Option<LineSpacingXml>,
    #[serde(rename = "textAlignment", alias = "w:textAlignment", default)]
    pub text_alignment: Option<XmlValueAttr>,
    #[serde(rename = "adjustRightInd", alias = "w:adjustRightInd", default)]
    pub adjust_right_ind: Option<XmlValueAttr>,
    #[serde(rename = "outlineLvl", alias = "w:outlineLvl", default)]
    pub outline_lvl: Option<XmlValueAttr>,
    #[serde(rename = "snapToGrid", alias = "w:snapToGrid", default)]
    pub snap_to_grid: Option<XmlOnOffAttr>,
    #[serde(rename = "keepNext", alias = "w:keepNext", default)]
    pub keep_next: Option<XmlOnOffAttr>,
    #[serde(rename = "keepLines", alias = "w:keepLines", default)]
    pub keep_lines: Option<XmlOnOffAttr>,
    #[serde(rename = "pageBreakBefore", alias = "w:pageBreakBefore", default)]
    pub page_break_before: Option<XmlOnOffAttr>,
    #[serde(rename = "widowControl", alias = "w:widowControl", default)]
    pub widow_control: Option<XmlOnOffAttr>,
    #[serde(rename = "divId", alias = "w:divId", default)]
    pub div_id: Option<XmlValueAttr>,
}

#[derive(Debug, Deserialize, Default)]
struct StyleXml {
    #[serde(rename = "@styleId", alias = "@w:styleId", default)]
    style_id: Option<String>,
    #[serde(rename = "@type", alias = "@w:type", default)]
    style_type: Option<String>,
    #[serde(rename = "name", alias = "w:name", default)]
    name: Option<XmlValueAttr>,
    #[serde(rename = "basedOn", alias = "w:basedOn", default)]
    based_on: Option<XmlValueAttr>,
    #[serde(rename = "next", alias = "w:next", default)]
    next: Option<XmlValueAttr>,
    #[serde(rename = "link", alias = "w:link", default)]
    link: Option<XmlValueAttr>,
    #[serde(rename = "rPr", alias = "w:rPr", default)]
    run_property: Option<RunPropertyXml>,
    #[serde(rename = "pPr", alias = "w:pPr", default)]
    paragraph_property: Option<ParagraphPropertyXml>,
}

// ============================================================================
// Parsing Helper Functions
// ============================================================================

fn parse_on_off(v: Option<&str>) -> bool {
    !matches!(
        v.map(|x| x.trim().to_ascii_lowercase()),
        Some(ref s) if s == "0" || s == "false"
    )
}

fn parse_usize(raw: Option<String>) -> Option<usize> {
    raw.and_then(|v| {
        v.parse::<usize>()
            .ok()
            .or_else(|| v.parse::<f32>().ok().map(|f| f as usize))
    })
}

fn parse_i32(raw: Option<String>) -> Option<i32> {
    raw.and_then(|v| {
        v.parse::<i32>()
            .ok()
            .or_else(|| v.parse::<f64>().ok().map(|f| f as i32))
    })
}

fn parse_u32(raw: Option<String>) -> Option<u32> {
    raw.and_then(|v| v.parse::<u32>().ok())
}

pub(crate) fn parse_run_property_xml(xml: Option<RunPropertyXml>) -> RunProperty {
    let Some(xml) = xml else {
        return RunProperty::new();
    };

    let mut rp = RunProperty::new();
    if let Some(v) = xml.style.and_then(|v| v.val) {
        rp = rp.style(&v);
    }
    if let Some(v) = parse_usize(xml.size.and_then(|v| v.val)) {
        rp = rp.size(v);
    }
    if let Some(v) = xml.color.and_then(|v| v.val) {
        rp = rp.color(v);
    }
    if let Some(v) = xml.highlight.and_then(|v| v.val) {
        rp = rp.highlight(v);
    }
    if let Some(v) = parse_i32(xml.spacing.and_then(|v| v.val)) {
        rp = rp.spacing(v);
    }
    if let Some(v) = xml.underline.and_then(|v| v.val) {
        rp = rp.underline(v);
    }
    if let Some(v) = xml.bold {
        rp = if parse_on_off(v.val.as_deref()) {
            rp.bold()
        } else {
            rp.disable_bold()
        };
    }
    if let Some(v) = xml.italic {
        rp = if parse_on_off(v.val.as_deref()) {
            rp.italic()
        } else {
            rp.disable_italic()
        };
    }
    if let Some(v) = xml.strike {
        rp = if parse_on_off(v.val.as_deref()) {
            rp.strike()
        } else {
            rp.disable_strike()
        };
    }
    if let Some(v) = xml.dstrike {
        rp = if parse_on_off(v.val.as_deref()) {
            rp.dstrike()
        } else {
            rp.disable_dstrike()
        };
    }
    if xml.vanish.is_some() {
        rp = rp.vanish();
    }
    if xml.spec_vanish.is_some() {
        rp = rp.spec_vanish();
    }
    if let Some(fonts) = xml.fonts {
        let mut f = RunFonts::new();
        if let Some(v) = fonts.ascii {
            f = f.ascii(v);
        }
        if let Some(v) = fonts.east_asia {
            f = f.east_asia(v);
        }
        if let Some(v) = fonts.h_ansi {
            f = f.hi_ansi(v);
        }
        if let Some(v) = fonts.cs {
            f = f.cs(v);
        }
        if let Some(v) = fonts.ascii_theme {
            f = f.ascii_theme(v);
        }
        if let Some(v) = fonts.east_asia_theme {
            f = f.east_asia_theme(v);
        }
        if let Some(v) = fonts.h_ansi_theme {
            f = f.hi_ansi_theme(v);
        }
        if let Some(v) = fonts.cs_theme {
            f = f.cs_theme(v);
        }
        if let Some(v) = fonts.hint {
            f = f.hint(v);
        }
        rp = rp.fonts(f);
    }

    rp
}

pub(crate) fn parse_paragraph_property_xml(xml: Option<ParagraphPropertyXml>) -> ParagraphProperty {
    let Some(xml) = xml else {
        return ParagraphProperty::new();
    };

    let mut p = ParagraphProperty::new();
    if let Some(v) = xml.style.and_then(|v| v.val) {
        p = p.style(&v);
    }
    if let Some(v) = xml.alignment.and_then(|v| v.val) {
        if let Ok(alignment) = AlignmentType::from_str(&v) {
            p = p.align(alignment);
        }
    }
    if let Some(v) = xml.text_alignment.and_then(|v| v.val) {
        if let Ok(text_alignment) = TextAlignmentType::from_str(&v) {
            p = p.text_alignment(text_alignment);
        }
    }
    if let Some(v) = parse_i32(xml.adjust_right_ind.and_then(|v| v.val)) {
        p = p.adjust_right_ind(v as isize);
    }
    if let Some(v) = parse_usize(xml.outline_lvl.and_then(|v| v.val)) {
        p = p.outline_lvl(v);
    }
    if let Some(v) = xml.snap_to_grid {
        p.snap_to_grid = Some(parse_on_off(v.val.as_deref()));
    }
    if let Some(v) = xml.keep_next {
        if parse_on_off(v.val.as_deref()) {
            p.keep_next = Some(true);
        }
    }
    if let Some(v) = xml.keep_lines {
        if parse_on_off(v.val.as_deref()) {
            p.keep_lines = Some(true);
        }
    }
    if let Some(v) = xml.page_break_before {
        if parse_on_off(v.val.as_deref()) {
            p.page_break_before = Some(true);
        }
    }
    if let Some(v) = xml.widow_control {
        if parse_on_off(v.val.as_deref()) {
            p.widow_control = Some(true);
        }
    }
    if let Some(v) = xml.div_id.and_then(|v| v.val) {
        p.div_id = Some(v);
    }
    if let Some(ind) = xml.indent {
        let start = parse_i32(ind.left);
        let end = parse_i32(ind.right);
        let special = if let Some(v) = parse_i32(ind.hanging.clone()) {
            Some(SpecialIndentType::Hanging(v))
        } else {
            parse_i32(ind.first_line.clone()).map(SpecialIndentType::FirstLine)
        };
        let start_chars = parse_i32(ind.start_chars);
        p = p.indent(start, special, end, start_chars);
        if let Some(v) = parse_i32(ind.hanging_chars) {
            p = p.hanging_chars(v);
        }
        if let Some(v) = parse_i32(ind.first_line_chars) {
            p = p.first_line_chars(v);
        }
    }
    if let Some(sp) = xml.spacing {
        let mut ls = LineSpacing::new();
        let mut has_spacing = false;
        if let Some(v) = sp.line_rule {
            if let Ok(rule) = LineSpacingType::from_str(&v) {
                ls = ls.line_rule(rule);
                has_spacing = true;
            }
        }
        if let Some(v) = parse_u32(sp.before) {
            ls = ls.before(v);
            has_spacing = true;
        }
        if let Some(v) = parse_u32(sp.after) {
            ls = ls.after(v);
            has_spacing = true;
        }
        if let Some(v) = parse_u32(sp.before_lines) {
            ls = ls.before_lines(v);
            has_spacing = true;
        }
        if let Some(v) = parse_u32(sp.after_lines) {
            ls = ls.after_lines(v);
            has_spacing = true;
        }
        if let Some(v) = parse_i32(sp.line) {
            ls = ls.line(v);
            has_spacing = true;
        }
        if has_spacing {
            p = p.line_spacing(ls);
        }
    }
    if let Some(run_property) = xml.run_property {
        p.run_property = parse_run_property_xml(Some(run_property));
    }

    p
}

impl<'de> Deserialize<'de> for Style {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let xml = StyleXml::deserialize(deserializer)?;
        let style_id = xml.style_id.unwrap_or_default();
        let style_type = xml
            .style_type
            .as_deref()
            .and_then(|v| StyleType::from_str(v).ok())
            .unwrap_or(StyleType::Paragraph);

        let mut style = Style::new(style_id, style_type);
        if let Some(v) = xml.name.and_then(|v| v.val) {
            style = style.name(v);
        }
        if let Some(v) = xml.based_on.and_then(|v| v.val) {
            style = style.based_on(v);
        }
        if let Some(v) = xml.next.and_then(|v| v.val) {
            style = style.next(v);
        }
        if let Some(v) = xml.link.and_then(|v| v.val) {
            style = style.link(v);
        }
        style.run_property = parse_run_property_xml(xml.run_property);
        style.paragraph_property = parse_paragraph_property_xml(xml.paragraph_property);
        Ok(style)
    }
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Style {
    pub style_id: String,
    pub name: Name,
    pub style_type: StyleType,
    pub run_property: RunProperty,
    pub paragraph_property: ParagraphProperty,
    pub table_property: TableProperty,
    pub table_cell_property: TableCellProperty,
    pub based_on: Option<BasedOn>,
    pub next: Option<Next>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub link: Option<Link>,
}

impl Default for Style {
    fn default() -> Self {
        let name = Name::new("");
        let rpr = RunProperty::new();
        let ppr = ParagraphProperty::new();
        Style {
            style_id: "".to_owned(),
            style_type: StyleType::Paragraph,
            name,
            run_property: rpr,
            paragraph_property: ppr,
            table_property: TableProperty::new(),
            table_cell_property: TableCellProperty::new(),
            based_on: None,
            next: None,
            link: None,
        }
    }
}

impl Style {
    pub fn new(style_id: impl Into<String>, style_type: StyleType) -> Self {
        let default = Default::default();
        Style {
            style_id: escape(&style_id.into()),
            style_type,
            ..default
        }
    }

    pub fn name(mut self, name: impl Into<String>) -> Self {
        self.name = Name::new(name);
        self
    }

    pub fn based_on(mut self, base: impl Into<String>) -> Self {
        self.based_on = Some(BasedOn::new(base));
        self
    }

    pub fn next(mut self, next: impl Into<String>) -> Self {
        self.next = Some(Next::new(next));
        self
    }

    pub fn link(mut self, link: impl Into<String>) -> Self {
        self.link = Some(Link::new(link));
        self
    }

    pub fn size(mut self, size: usize) -> Self {
        self.run_property = self.run_property.size(size);
        self
    }

    pub fn color(mut self, color: impl Into<String>) -> Self {
        self.run_property = self.run_property.color(color);
        self
    }

    pub fn highlight(mut self, color: impl Into<String>) -> Self {
        self.run_property = self.run_property.highlight(color);
        self
    }

    pub fn bold(mut self) -> Self {
        self.run_property = self.run_property.bold();
        self
    }

    pub fn italic(mut self) -> Self {
        self.run_property = self.run_property.italic();
        self
    }

    pub fn underline(mut self, line_type: impl Into<String>) -> Self {
        self.run_property = self.run_property.underline(line_type);
        self
    }

    pub fn vanish(mut self) -> Self {
        self.run_property = self.run_property.vanish();
        self
    }

    pub fn text_border(mut self, b: TextBorder) -> Self {
        self.run_property = self.run_property.text_border(b);
        self
    }

    pub fn fonts(mut self, f: RunFonts) -> Self {
        self.run_property = self.run_property.fonts(f);
        self
    }

    pub fn align(mut self, alignment_type: AlignmentType) -> Self {
        self.paragraph_property = self.paragraph_property.align(alignment_type);
        self
    }

    pub fn text_alignment(mut self, alignment_type: TextAlignmentType) -> Self {
        self.paragraph_property = self.paragraph_property.text_alignment(alignment_type);
        self
    }

    pub fn snap_to_grid(mut self, v: bool) -> Self {
        self.paragraph_property = self.paragraph_property.snap_to_grid(v);
        self
    }

    pub fn line_spacing(mut self, spacing: LineSpacing) -> Self {
        self.paragraph_property = self.paragraph_property.line_spacing(spacing);
        self
    }

    pub fn indent(
        mut self,
        left: Option<i32>,
        special_indent: Option<SpecialIndentType>,
        end: Option<i32>,
        start_chars: Option<i32>,
    ) -> Self {
        self.paragraph_property =
            self.paragraph_property
                .indent(left, special_indent, end, start_chars);
        self
    }

    pub fn hanging_chars(mut self, chars: i32) -> Self {
        self.paragraph_property = self.paragraph_property.hanging_chars(chars);
        self
    }

    pub fn first_line_chars(mut self, chars: i32) -> Self {
        self.paragraph_property = self.paragraph_property.first_line_chars(chars);
        self
    }

    pub fn outline_lvl(mut self, l: usize) -> Self {
        self.paragraph_property = self.paragraph_property.outline_lvl(l);
        self
    }

    pub fn table_property(mut self, p: TableProperty) -> Self {
        self.table_property = p;
        self
    }

    pub fn table_indent(mut self, v: i32) -> Self {
        self.table_property = self.table_property.indent(v);
        self
    }

    pub fn table_align(mut self, v: TableAlignmentType) -> Self {
        self.table_property = self.table_property.align(v);
        self
    }

    pub fn style(mut self, s: impl Into<String>) -> Self {
        self.table_property = self.table_property.style(s);
        self
    }

    pub fn layout(mut self, t: TableLayoutType) -> Self {
        self.table_property = self.table_property.layout(t);
        self
    }

    pub fn width(mut self, w: usize, t: WidthType) -> Self {
        self.table_property = self.table_property.width(w, t);
        self
    }

    pub fn margins(mut self, margins: TableCellMargins) -> Self {
        self.table_property = self.table_property.set_margins(margins);
        self
    }

    pub fn set_borders(mut self, borders: TableBorders) -> Self {
        self.table_property = self.table_property.set_borders(borders);
        self
    }

    pub fn set_border(mut self, border: TableBorder) -> Self {
        self.table_property = self.table_property.set_border(border);
        self
    }

    pub fn clear_border(mut self, position: TableBorderPosition) -> Self {
        self.table_property = self.table_property.clear_border(position);
        self
    }

    pub fn clear_all_border(mut self) -> Self {
        self.table_property = self.table_property.clear_all_border();
        self
    }

    pub fn table_cell_property(mut self, p: TableCellProperty) -> Self {
        self.table_cell_property = p;
        self
    }

    // frameProperty
    pub fn wrap(mut self, wrap: impl Into<String>) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            wrap: Some(wrap.into()),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn v_anchor(mut self, anchor: impl Into<String>) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            v_anchor: Some(anchor.into()),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn h_anchor(mut self, anchor: impl Into<String>) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            h_anchor: Some(anchor.into()),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn h_rule(mut self, r: impl Into<String>) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            h_rule: Some(r.into()),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn x_align(mut self, align: impl Into<String>) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            x_align: Some(align.into()),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn y_align(mut self, align: impl Into<String>) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            y_align: Some(align.into()),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn h_space(mut self, x: i32) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            h_space: Some(x),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn v_space(mut self, x: i32) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            v_space: Some(x),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn frame_x(mut self, x: i32) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            x: Some(x),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn frame_y(mut self, y: i32) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            y: Some(y),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn frame_width(mut self, n: u32) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            w: Some(n),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }

    pub fn frame_height(mut self, n: u32) -> Self {
        self.paragraph_property.frame_property = Some(FrameProperty {
            h: Some(n),
            ..self.paragraph_property.frame_property.unwrap_or_default()
        });
        self
    }
}

impl BuildXML for Style {
    fn build_to<W: Write>(
        &self,
        stream: xml::writer::EventWriter<W>,
    ) -> xml::writer::Result<xml::writer::EventWriter<W>> {
        // Set "Normal" as default if you need change these values please fix it
        XMLBuilder::from(stream)
            .open_style(self.style_type, &self.style_id)?
            .add_child(&self.name)?
            .add_child(&self.run_property)?
            .add_child(&self.paragraph_property)?
            .apply_if(self.style_type == StyleType::Table, |b| {
                b.add_child(&self.table_cell_property)?
                    .add_child(&self.table_property)
            })?
            .add_optional_child(&self.next)?
            .add_optional_child(&self.link)?
            .add_child(&QFormat::new())?
            .add_optional_child(&self.based_on)?
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
    fn test_build() {
        let c = Style::new("Heading", StyleType::Paragraph).name("Heading1");
        let b = c.build();
        assert_eq!(
            str::from_utf8(&b).unwrap(),
            r#"<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading1" /><w:rPr /><w:pPr><w:rPr /></w:pPr><w:qFormat /></w:style>"#
        );
    }
}
