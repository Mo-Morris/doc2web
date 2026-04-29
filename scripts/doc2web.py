#!/usr/bin/env python3
"""Convert a Word document into a complete static website.

The converter intentionally avoids network and third-party package dependencies.
It parses .docx files directly from the OOXML package and optionally uses
LibreOffice/soffice only when an old .doc file needs conversion first.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import mimetypes
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def qn(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def attr(el: ET.Element, prefix: str, name: str) -> str | None:
    return el.attrib.get(qn(prefix, name))


@dataclass
class RenderContext:
    zipf: zipfile.ZipFile
    out_dir: Path
    rels: dict[str, dict[str, str]]
    styles: dict[str, dict[str, str]]
    base_dir: str = "word"
    media_seen: dict[str, str] = field(default_factory=dict)
    media_count: int = 0


@dataclass
class Block:
    kind: str
    html: str
    text: str
    level: int = 0
    title: str = ""
    anchor: str = ""


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def read_xml(zipf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        with zipf.open(name) as fh:
            return ET.fromstring(fh.read())
    except KeyError:
        return None


def parse_rels(zipf: zipfile.ZipFile, rel_path: str) -> dict[str, dict[str, str]]:
    root = read_xml(zipf, rel_path)
    if root is None:
        return {}
    rels: dict[str, dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id")
        if rid:
            rels[rid] = {
                "target": rel.attrib.get("Target", ""),
                "type": rel.attrib.get("Type", ""),
                "mode": rel.attrib.get("TargetMode", ""),
            }
    return rels


def parse_styles(zipf: zipfile.ZipFile) -> dict[str, dict[str, str]]:
    root = read_xml(zipf, "word/styles.xml")
    styles: dict[str, dict[str, str]] = {}
    if root is None:
        return styles
    for style in root.findall("w:style", NS):
        style_id = attr(style, "w", "styleId")
        if not style_id:
            continue
        name_el = style.find("w:name", NS)
        name = attr(name_el, "w", "val") if name_el is not None else style_id
        outline = ""
        outline_el = style.find("w:pPr/w:outlineLvl", NS)
        if outline_el is not None:
            outline = attr(outline_el, "w", "val") or ""
        styles[style_id] = {"name": name or style_id, "outline": outline}
    return styles


def heading_level(p: ET.Element, styles: dict[str, dict[str, str]]) -> int:
    style_el = p.find("w:pPr/w:pStyle", NS)
    if style_el is None:
        return 0
    style_id = attr(style_el, "w", "val") or ""
    info = styles.get(style_id, {})
    name = (info.get("name") or style_id).lower().strip()
    outline = info.get("outline", "")
    if outline.isdigit():
        return min(int(outline) + 1, 6)
    match = re.search(r"(heading|标题|titre|überschrift|rubrik)\s*([1-6])", name)
    if match:
        return int(match.group(2))
    return 0


def paragraph_class(p: ET.Element, styles: dict[str, dict[str, str]]) -> str:
    style_el = p.find("w:pPr/w:pStyle", NS)
    if style_el is None:
        return ""
    style_id = attr(style_el, "w", "val") or ""
    name = styles.get(style_id, {}).get("name", style_id).lower()
    if "quote" in name or "引用" in name:
        return "quote"
    if "caption" in name or "题注" in name:
        return "caption"
    return ""


def slugify(text: str, used: set[str]) -> str:
    value = re.sub(r"[^\w\u4e00-\u9fff]+", "-", text.lower(), flags=re.UNICODE).strip("-")
    value = value[:80] or "section"
    base = value
    index = 2
    while value in used:
        value = f"{base}-{index}"
        index += 1
    used.add(value)
    return value


def relationship_target(ctx: RenderContext, rid: str) -> dict[str, str] | None:
    return ctx.rels.get(rid)


def rels_for_part(zipf: zipfile.ZipFile, part: str) -> dict[str, dict[str, str]]:
    path = Path(part)
    rel_path = f"{path.parent}/_rels/{path.name}.rels"
    return parse_rels(zipf, rel_path)


def extract_media(ctx: RenderContext, rid: str) -> str:
    rel = relationship_target(ctx, rid)
    if not rel:
        return ""
    target = rel.get("target", "")
    if rel.get("mode") == "External":
        return target
    if target.startswith("/"):
        source = target.lstrip("/")
    else:
        source = posixpath.normpath(posixpath.join(ctx.base_dir, target))
    if source in ctx.media_seen:
        return ctx.media_seen[source]
    try:
        data = ctx.zipf.read(source)
    except KeyError:
        return ""
    digest = hashlib.sha1(data).hexdigest()[:10]
    suffix = Path(source).suffix or mimetypes.guess_extension(mimetypes.guess_type(source)[0] or "") or ".bin"
    ctx.media_count += 1
    filename = f"media/{ctx.media_count:03d}-{digest}{suffix}"
    destination = ctx.out_dir / "assets" / filename
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(data)
    public_path = f"assets/{filename}"
    ctx.media_seen[source] = public_path
    return public_path


def image_html(ctx: RenderContext, container: ET.Element) -> tuple[str, str]:
    parts: list[str] = []
    labels: list[str] = []
    for blip in container.findall(".//a:blip", NS):
        rid = attr(blip, "r", "embed") or attr(blip, "r", "link")
        if not rid:
            continue
        src = extract_media(ctx, rid)
        if not src:
            continue
        alt = ""
        doc_pr = container.find(".//wp:docPr", NS)
        if doc_pr is not None:
            alt = doc_pr.attrib.get("descr") or doc_pr.attrib.get("name") or ""
        labels.append(alt or "Image")
        parts.append(
            f'<figure class="doc-figure"><img src="{html.escape(src)}" alt="{html.escape(alt)}"></figure>'
        )
    for image_data in container.findall(".//v:imagedata", NS):
        rid = attr(image_data, "r", "id")
        if not rid:
            continue
        src = extract_media(ctx, rid)
        if src:
            labels.append("Image")
            parts.append(f'<figure class="doc-figure"><img src="{html.escape(src)}" alt=""></figure>')
    return "".join(parts), " ".join(labels)


def run_html(ctx: RenderContext, run: ET.Element) -> tuple[str, str]:
    rpr = run.find("w:rPr", NS)
    is_bold = rpr is not None and rpr.find("w:b", NS) is not None
    is_italic = rpr is not None and rpr.find("w:i", NS) is not None
    is_underlined = rpr is not None and rpr.find("w:u", NS) is not None
    pieces: list[str] = []
    plain: list[str] = []
    for child in run:
        name = local_name(child.tag)
        if name == "t":
            value = child.text or ""
            pieces.append(html.escape(value))
            plain.append(value)
        elif name == "tab":
            pieces.append(" ")
            plain.append(" ")
        elif name in {"br", "cr"}:
            pieces.append("<br>")
            plain.append("\n")
        elif name in {"drawing", "pict"}:
            rendered, label = image_html(ctx, child)
            pieces.append(rendered)
            plain.append(label)
    rendered = "".join(pieces)
    if not rendered:
        return "", ""
    if is_underlined:
        rendered = f"<u>{rendered}</u>"
    if is_italic:
        rendered = f"<em>{rendered}</em>"
    if is_bold:
        rendered = f"<strong>{rendered}</strong>"
    return rendered, "".join(plain)


def inline_html(ctx: RenderContext, parent: ET.Element) -> tuple[str, str]:
    pieces: list[str] = []
    plain: list[str] = []
    for child in parent:
        name = local_name(child.tag)
        if name == "r":
            rendered, text = run_html(ctx, child)
            pieces.append(rendered)
            plain.append(text)
        elif name == "hyperlink":
            inner_html, inner_text = inline_html(ctx, child)
            rid = attr(child, "r", "id")
            anchor = attr(child, "w", "anchor")
            href = ""
            if rid and relationship_target(ctx, rid):
                href = relationship_target(ctx, rid).get("target", "")
            elif anchor:
                href = f"#{anchor}"
            if href:
                pieces.append(f'<a href="{html.escape(href)}">{inner_html}</a>')
            else:
                pieces.append(inner_html)
            plain.append(inner_text)
        elif name in {"drawing", "pict"}:
            rendered, label = image_html(ctx, child)
            pieces.append(rendered)
            plain.append(label)
    return "".join(pieces), "".join(plain)


def paragraph_block(ctx: RenderContext, p: ET.Element, used: set[str], styles: dict[str, dict[str, str]]) -> Block | None:
    body, text = inline_html(ctx, p)
    text = re.sub(r"\s+", " ", text).strip()
    if not body.strip() and not text:
        return None
    level = heading_level(p, styles)
    if level:
        anchor = slugify(text, used)
        return Block("heading", f'<h{level} id="{anchor}">{body}</h{level}>', text, level, text, anchor)
    css = paragraph_class(p, styles)
    if css == "quote":
        return Block("paragraph", f'<blockquote>{body}</blockquote>', text)
    if css == "caption":
        return Block("paragraph", f'<p class="caption">{body}</p>', text)
    return Block("paragraph", f"<p>{body}</p>", text)


def table_block(ctx: RenderContext, table: ET.Element, used: set[str], styles: dict[str, dict[str, str]]) -> Block:
    rows: list[str] = []
    text_parts: list[str] = []
    for tr in table.findall("w:tr", NS):
        cells: list[str] = []
        row_text: list[str] = []
        for tc in tr.findall("w:tc", NS):
            inner: list[str] = []
            cell_text: list[str] = []
            for child in tc:
                name = local_name(child.tag)
                if name == "p":
                    block = paragraph_block(ctx, child, used, styles)
                    if block:
                        inner.append(block.html)
                        cell_text.append(block.text)
                elif name == "tbl":
                    block = table_block(ctx, child, used, styles)
                    inner.append(block.html)
                    cell_text.append(block.text)
            cells.append(f"<td>{''.join(inner)}</td>")
            row_text.append(" ".join(cell_text))
        rows.append(f"<tr>{''.join(cells)}</tr>")
        text_parts.append(" | ".join(row_text))
    table_html = '<div class="table-wrap"><table>' + "".join(rows) + "</table></div>"
    return Block("table", table_html, "\n".join(text_parts))


def blocks_from_part(ctx: RenderContext, part: str, used: set[str]) -> list[Block]:
    root = read_xml(ctx.zipf, part)
    if root is None:
        return []
    old_rels = ctx.rels
    old_base_dir = ctx.base_dir
    part_rels = rels_for_part(ctx.zipf, part)
    if part_rels:
        ctx.rels = part_rels
    ctx.base_dir = str(Path(part).parent).replace("\\", "/")
    blocks: list[Block] = []
    body = root.find("w:body", NS)
    if body is None:
        body = root
    try:
        for child in body:
            name = local_name(child.tag)
            if name == "p":
                block = paragraph_block(ctx, child, used, ctx.styles)
                if block:
                    blocks.append(block)
            elif name == "tbl":
                blocks.append(table_block(ctx, child, used, ctx.styles))
        return blocks
    finally:
        ctx.rels = old_rels
        ctx.base_dir = old_base_dir


def note_blocks(ctx: RenderContext, name: str, label: str, used: set[str]) -> list[Block]:
    root = read_xml(ctx.zipf, name)
    if root is None:
        return []
    old_rels = ctx.rels
    old_base_dir = ctx.base_dir
    part_rels = rels_for_part(ctx.zipf, name)
    if part_rels:
        ctx.rels = part_rels
    ctx.base_dir = str(Path(name).parent).replace("\\", "/")
    blocks: list[Block] = []
    try:
        note_items = [el for el in root if local_name(el.tag).lower().endswith("note")]
        if not note_items:
            return []
        anchor = slugify(label, used)
        blocks.append(Block("heading", f'<h2 id="{anchor}">{label}</h2>', label, 2, label, anchor))
        for index, note in enumerate(note_items, 1):
            note_type = attr(note, "w", "type")
            if note_type in {"separator", "continuationSeparator"}:
                continue
            inner: list[str] = []
            text_parts: list[str] = []
            for p in note.findall("w:p", NS):
                block = paragraph_block(ctx, p, used, ctx.styles)
                if block:
                    inner.append(block.html)
                    text_parts.append(block.text)
            if inner:
                blocks.append(Block("note", f'<aside class="note"><span>{index}</span>{"".join(inner)}</aside>', " ".join(text_parts)))
        return blocks
    finally:
        ctx.rels = old_rels
        ctx.base_dir = old_base_dir


def header_footer_blocks(ctx: RenderContext, used: set[str]) -> list[Block]:
    names = sorted(
        name for name in ctx.zipf.namelist()
        if re.match(r"word/(header|footer)\d+\.xml$", name)
    )
    blocks: list[Block] = []
    for name in names:
        part_blocks = blocks_from_part(ctx, name, used)
        if not part_blocks:
            continue
        title = "Headers And Footers" if not blocks else ""
        if title:
            anchor = slugify(title, used)
            blocks.append(Block("heading", f'<h2 id="{anchor}">{title}</h2>', title, 2, title, anchor))
        blocks.extend(part_blocks)
    return blocks


def build_toc(blocks: Iterable[Block]) -> list[dict[str, str | int]]:
    return [
        {"level": block.level, "title": block.title, "anchor": block.anchor}
        for block in blocks
        if block.kind == "heading" and block.anchor
    ]


def build_sections(blocks: list[Block]) -> list[dict[str, str]]:
    sections: list[dict[str, str]] = []
    current = {"title": "Introduction", "anchor": "top", "text": []}
    for block in blocks:
        if block.kind == "heading" and block.level <= 2:
            if current["text"]:
                sections.append({"title": current["title"], "anchor": current["anchor"], "text": " ".join(current["text"])})
            current = {"title": block.title, "anchor": block.anchor, "text": []}
        else:
            current["text"].append(block.text)
    if current["text"]:
        sections.append({"title": current["title"], "anchor": current["anchor"], "text": " ".join(current["text"])})
    return sections


def compact_text(value: str, limit: int = 180) -> str:
    value = re.sub(r"\s+", " ", value).strip()
    if len(value) <= limit:
        return value
    clipped = value[:limit].rsplit(" ", 1)[0].strip()
    return f"{clipped or value[:limit].strip()}..."


def reading_time(word_count: int) -> int:
    return max(1, round(word_count / 420))


def page_language(title: str, blocks: list[Block]) -> str:
    sample = title + " " + " ".join(block.text for block in blocks[:16])
    cjk = len(re.findall(r"[\u4e00-\u9fff]", sample))
    return "zh-CN" if cjk > 12 else "en"


def first_image_src(blocks: list[Block]) -> str:
    for block in blocks:
        match = re.search(r'<img\s+[^>]*src="([^"]+)"', block.html)
        if match:
            return match.group(1)
    return ""


def install_hero_image(image_path: Path | None, out_dir: Path) -> str:
    if image_path is None:
        return ""
    source = image_path.expanduser().resolve()
    if not source.exists():
        raise SystemExit(f"Hero image not found: {source}")
    suffix = source.suffix or mimetypes.guess_extension(mimetypes.guess_type(source.name)[0] or "") or ".png"
    data = source.read_bytes()
    digest = hashlib.sha1(data).hexdigest()[:10]
    destination = out_dir / "assets" / f"hero-{digest}{suffix}"
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source != destination:
        destination.write_bytes(data)
    return f"assets/{destination.name}"


def sections_with_text(sections: list[dict[str, str]]) -> list[dict[str, str]]:
    usable = [section for section in sections if section["text"].strip()]
    return usable or sections[:]


def top_sections(sections: list[dict[str, str]], count: int = 6) -> list[dict[str, str]]:
    usable = sections_with_text(sections)
    ranked = sorted(usable, key=lambda item: len(item["text"]), reverse=True)
    selected: list[dict[str, str]] = []
    seen: set[str] = set()
    for section in ranked:
        key = section["anchor"]
        if key not in seen:
            selected.append(section)
            seen.add(key)
        if len(selected) >= count:
            break
    return selected


def section_matching(sections: list[dict[str, str]], patterns: list[str]) -> dict[str, str] | None:
    regex = re.compile("|".join(patterns), re.IGNORECASE)
    for section in sections_with_text(sections):
        haystack = f'{section["title"]} {section["text"]}'
        if regex.search(haystack):
            return section
    return None


def ui_copy(lang: str) -> dict[str, str]:
    if lang == "zh-CN":
        return {
            "nav_insights": "策展维度",
            "nav_themes": "精选主题",
            "nav_document": "完整文档",
            "eyebrow": "官网级文档体验",
            "fallback_lead": "基于源文档生成的结构化官网站点。",
            "sections": "个章节",
            "words": "个词",
            "read_minutes": "分钟阅读",
            "insights_label": "策展维度",
            "insights_title": "不是文档搬运，而是战略入口。",
            "insights_text": "内容先被重组为高信号入口，帮助读者先理解主线、证据与下一步，再进入完整文档。",
            "themes_label": "精选主题",
            "themes_title": "从多个角度进入重点内容。",
            "themes_text": "这些卡片来自信息密度最高的章节，并链接回下方保留的源文档内容。",
            "chapter": "章节",
            "read_section": "阅读章节",
            "document_label": "完整来源",
            "document_title": "完整文档，可搜索、可导航。",
            "document_text": "解析出的标题、段落、表格、图片、注释、页眉与页脚都会保留，便于核验与深度阅读。",
            "search": "搜索文档",
            "toc": "目录",
            "footer_prefix": "由源文档生成",
            "footer_suffix": "静态 Doc2Web 站点。",
        }
    return {
        "nav_insights": "Insights",
        "nav_themes": "Themes",
        "nav_document": "Document",
        "eyebrow": "Official Document Experience",
        "fallback_lead": "A polished, structured website generated from the source document.",
        "sections": "sections",
        "words": "words",
        "read_minutes": "min read",
        "insights_label": "Curated Dimensions",
        "insights_title": "Not a document dump. A strategic front door.",
        "insights_text": "The source is reorganized into high-signal entry points first, so readers can understand the story, evidence, and next moves before entering the full document.",
        "themes_label": "Featured Themes",
        "themes_title": "Explore the strongest sections from multiple angles.",
        "themes_text": "These cards are drawn from the densest sections and link back to the preserved source content below.",
        "chapter": "Chapter",
        "read_section": "Read section",
        "document_label": "Complete Source",
        "document_title": "Full document, searchable and structured.",
        "document_text": "Every parsed heading, paragraph, table, image, note, header, and footer remains available for verification and detailed reading.",
        "search": "Search this document",
        "toc": "Table of contents",
        "footer_prefix": "Generated from",
        "footer_suffix": "as a static Doc2Web site.",
    }


def insight_cards(sections: list[dict[str, str]], toc: list[dict[str, str | int]], word_count: int, lang: str) -> list[dict[str, str]]:
    themes = [str(item["title"]) for item in toc if int(item["level"]) <= 2][:5]
    dense = top_sections(sections, 1)
    evidence = section_matching(sections, [r"\d+(\.\d+)?%?", "data", "metric", "figure", "table", "数据", "指标", "规模", "增长"])
    action = section_matching(sections, ["action", "plan", "roadmap", "next", "risk", "recommend", "建议", "行动", "计划", "路线", "风险"])
    if lang == "zh-CN":
        labels = [
            ("叙事", "核心主题", "从文档标题与正文自动提炼主要叙事线。"),
            ("信号", dense[0]["title"] if dense else "关键信息", "优先呈现正文中信息密度最高的部分。"),
            ("证据", evidence["title"] if evidence else "数据与依据", f"全文约 {word_count:,} 个词，包含 {len(toc)} 个可导航章节。"),
            ("方向", action["title"] if action else "行动线索", "从计划、风险、建议与后续步骤等维度组织阅读入口。"),
        ]
    else:
        labels = [
            ("Narrative", "Core themes", "The main storyline inferred from headings and source text."),
            ("Signal", dense[0]["title"] if dense else "Key information", "Prioritizes the densest part of the source document."),
            ("Evidence", evidence["title"] if evidence else "Evidence base", f"About {word_count:,} words across {len(toc)} navigable sections."),
            ("Direction", action["title"] if action else "Action cues", "Organizes plans, risks, recommendations, and next steps into a clear entry point."),
        ]
    cards = [
        {
            "kicker": labels[0][0],
            "title": labels[0][1],
            "text": " / ".join(themes) if themes else labels[0][2],
        },
        {
            "kicker": labels[1][0],
            "title": labels[1][1],
            "text": compact_text(dense[0]["text"] if dense else "", 150) or labels[1][2],
        },
        {
            "kicker": labels[2][0],
            "title": labels[2][1],
            "text": compact_text(evidence["text"] if evidence else labels[2][2], 150),
        },
        {
            "kicker": labels[3][0],
            "title": labels[3][1],
            "text": compact_text(action["text"] if action else labels[3][2], 150),
        },
    ]
    return cards


def css() -> str:
    return r"""
:root {
  --ink: #101315;
  --muted: #5d666f;
  --paper: #f6f2ea;
  --panel: #ffffff;
  --line: rgba(16, 19, 21, 0.12);
  --accent: #ff5f45;
  --accent-2: #00a6a6;
  --accent-3: #c6ff4d;
  --dark: #14181b;
  --shadow: 0 22px 70px rgba(16, 19, 21, 0.14);
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
  margin: 0;
  color: var(--ink);
  background: var(--paper);
  font-family: Inter, Avenir Next, Helvetica Neue, Arial, sans-serif;
}
.progress {
  position: fixed;
  inset: 0 auto auto 0;
  width: 0;
  height: 4px;
  z-index: 20;
  background: linear-gradient(90deg, var(--accent), var(--accent-2));
}
.topbar {
  position: sticky;
  top: 0;
  z-index: 15;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 18px;
  padding: 14px clamp(18px, 4vw, 56px);
  border-bottom: 1px solid rgba(255,255,255,0.12);
  color: #fff;
  background: rgba(20, 24, 27, 0.88);
  backdrop-filter: blur(18px);
}
.brand {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  color: #fff;
  font-size: 0.76rem;
  font-weight: 900;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}
.brand::before {
  width: 10px;
  height: 28px;
  content: "";
  background: linear-gradient(180deg, var(--accent), var(--accent-2) 60%, var(--accent-3));
}
.toplinks {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}
.toplinks a {
  padding: 8px 10px;
  color: rgba(255,255,255,0.78);
  text-decoration: none;
  font-size: 0.85rem;
  font-weight: 800;
}
.toplinks a:hover, .toplinks a:focus-visible { color: #fff; outline: 2px solid var(--accent-3); outline-offset: 2px; }
.hero {
  min-height: 86vh;
  display: grid;
  align-items: end;
  padding: clamp(72px, 10vw, 128px) clamp(18px, 5vw, 72px) clamp(34px, 6vw, 72px);
  color: #fff;
  background:
    linear-gradient(120deg, rgba(20,24,27,0.96) 0%, rgba(20,24,27,0.82) 48%, rgba(20,24,27,0.44) 100%),
    var(--hero-image, linear-gradient(135deg, #14181b 0%, #364049 100%));
  background-size: cover;
  background-position: center;
}
.hero-inner {
  width: min(1180px, 100%);
}
.eyebrow {
  width: fit-content;
  margin-bottom: 18px;
  padding: 8px 10px;
  color: var(--accent-3);
  border: 1px solid rgba(198,255,77,0.36);
  background: rgba(198,255,77,0.08);
  font-size: 0.76rem;
  font-weight: 900;
  letter-spacing: 0.14em;
  text-transform: uppercase;
}
h1, h2, h3, h4, h5, h6 {
  margin: 1.3em 0 0.46em;
  line-height: 1.08;
  letter-spacing: 0;
  text-wrap: balance;
}
h1 {
  max-width: 980px;
  margin: 0;
  font-size: clamp(3.2rem, 8vw, 7.8rem);
  font-weight: 950;
}
.dek {
  max-width: 760px;
  margin: 24px 0 0;
  color: rgba(255,255,255,0.78);
  font-size: clamp(1.1rem, 2vw, 1.5rem);
  line-height: 1.55;
}
.meta {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 28px;
}
.meta span {
  padding: 9px 12px;
  border: 1px solid rgba(255,255,255,0.18);
  background: rgba(255,255,255,0.08);
  color: rgba(255,255,255,0.82);
  font-size: 0.82rem;
  font-weight: 850;
}
.band {
  padding: clamp(44px, 7vw, 92px) clamp(18px, 5vw, 72px);
}
.band.dark {
  color: #fff;
  background: var(--dark);
}
.section-head {
  max-width: 900px;
  margin: 0 auto 30px;
}
.section-head .label {
  color: var(--accent-2);
  font-size: 0.76rem;
  font-weight: 950;
  letter-spacing: 0.14em;
  text-transform: uppercase;
}
.dark .section-head .label { color: var(--accent-3); }
h2 {
  margin-top: 10px;
  font-size: clamp(2rem, 4vw, 4.6rem);
  font-weight: 950;
}
.section-head p {
  max-width: 720px;
  color: var(--muted);
  font-size: 1.06rem;
  line-height: 1.7;
}
.dark .section-head p { color: rgba(255,255,255,0.68); }
.insights {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 12px;
  width: min(1180px, 100%);
  margin: 0 auto;
}
.insight, .topic-card {
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
  box-shadow: var(--shadow);
}
.insight {
  min-height: 260px;
  padding: 22px;
}
.insight:nth-child(2) { border-top: 5px solid var(--accent); }
.insight:nth-child(3) { border-top: 5px solid var(--accent-2); }
.insight:nth-child(4) { border-top: 5px solid var(--accent-3); }
.kicker {
  color: var(--accent-2);
  font-size: 0.72rem;
  font-weight: 950;
  letter-spacing: 0.14em;
  text-transform: uppercase;
}
.insight h3, .topic-card h3 {
  margin-top: 16px;
  font-size: 1.35rem;
  line-height: 1.15;
}
.insight p, .topic-card p {
  color: var(--muted);
  line-height: 1.65;
}
.topic-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 14px;
  width: min(1180px, 100%);
  margin: 0 auto;
}
.topic-card {
  display: flex;
  min-height: 250px;
  flex-direction: column;
  justify-content: space-between;
  padding: 22px;
  color: #fff;
  background: linear-gradient(145deg, #1c2226 0%, #2a302e 100%);
  box-shadow: none;
}
.topic-card:nth-child(2n) { background: linear-gradient(145deg, #163c3f 0%, #24302d 100%); }
.topic-card:nth-child(3n) { background: linear-gradient(145deg, #40241f 0%, #202a2b 100%); }
.topic-card p { color: rgba(255,255,255,0.70); }
.topic-card a {
  width: fit-content;
  margin-top: 18px;
  color: var(--accent-3);
  font-weight: 900;
  text-decoration: none;
}
.shell {
  display: grid;
  grid-template-columns: minmax(230px, 310px) minmax(0, 1fr);
  gap: clamp(1rem, 3vw, 3rem);
  width: min(1480px, calc(100% - 32px));
  margin: 0 auto;
  padding: 0 0 72px;
}
.sidebar {
  position: sticky;
  top: 24px;
  align-self: start;
  max-height: calc(100vh - 48px);
  overflow: auto;
  padding: 20px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.88);
  box-shadow: var(--shadow);
  backdrop-filter: blur(18px);
}
.search {
  width: 100%;
  margin-bottom: 16px;
  padding: 12px 14px;
  border: 1px solid var(--line);
  border-radius: 8px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.72);
  font: 700 0.95rem Inter, Avenir Next, Helvetica Neue, Arial, sans-serif;
}
.results {
  display: none;
  margin-bottom: 14px;
  padding: 10px;
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.62);
}
.results a, .toc a {
  display: block;
  color: var(--ink);
  text-decoration: none;
}
.results a {
  padding: 8px 10px;
  border-radius: 6px;
  font-size: 0.9rem;
}
.results a:hover, .toc a:hover { background: rgba(0, 166, 166, 0.10); }
.toc {
  display: grid;
  gap: 3px;
  padding-top: 10px;
  border-top: 1px solid var(--line);
}
.toc a {
  padding: 7px 9px;
  border-radius: 6px;
  color: var(--muted);
  font-size: 0.88rem;
  line-height: 1.3;
}
.toc .level-1 { margin-left: 0; color: var(--ink); font-weight: 800; }
.toc .level-2 { margin-left: 10px; }
.toc .level-3 { margin-left: 20px; font-size: 0.82rem; }
.toc .level-4, .toc .level-5, .toc .level-6 { margin-left: 30px; font-size: 0.78rem; }
h3 { font-size: clamp(1.55rem, 2.7vw, 2.4rem); }
h4 { font-size: 1.35rem; }
.content {
  padding: clamp(1.2rem, 5vw, 4.5rem);
  border: 1px solid var(--line);
  border-radius: 8px;
  background: var(--panel);
  box-shadow: var(--shadow);
  backdrop-filter: blur(14px);
  font-family: Charter, "Iowan Old Style", "Palatino Linotype", Georgia, serif;
}
.content p, .content li, .content td {
  font-size: clamp(1.04rem, 1.4vw, 1.18rem);
  line-height: 1.78;
}
.content p {
  margin: 0.75em 0;
}
.content a {
  color: var(--accent-2);
  text-decoration-thickness: 0.08em;
  text-underline-offset: 0.18em;
}
blockquote {
  margin: 1.5rem 0;
  padding: 1rem 1.3rem;
  border-left: 5px solid var(--accent);
  border-radius: 8px;
  background: rgba(180, 95, 42, 0.08);
  color: #4b453c;
}
.caption {
  color: var(--muted);
  font-size: 0.95rem !important;
  font-style: italic;
}
.doc-figure {
  margin: 2rem 0;
  padding: 12px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255,255,255,0.68);
}
.doc-figure img {
  display: block;
  max-width: 100%;
  height: auto;
  margin: 0 auto;
  border-radius: 6px;
}
.table-wrap {
  width: 100%;
  margin: 1.6rem 0;
  overflow-x: auto;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255,255,255,0.76);
}
table {
  width: 100%;
  min-width: 720px;
  border-collapse: collapse;
}
td {
  min-width: 160px;
  padding: 12px 14px;
  border: 1px solid rgba(23, 33, 27, 0.11);
  vertical-align: top;
}
td p { margin: 0.25rem 0 !important; }
.note {
  display: grid;
  grid-template-columns: 36px 1fr;
  gap: 12px;
  margin: 0.75rem 0;
  padding: 12px;
  border-radius: 8px;
  background: rgba(28,111,98,0.08);
}
.note span {
  display: grid;
  place-items: center;
  width: 28px;
  height: 28px;
  border-radius: 50%;
  color: white;
  background: var(--accent-2);
  font: 800 0.8rem Inter, Avenir Next, Helvetica Neue, Arial, sans-serif;
}
.footer {
  padding: 28px clamp(18px, 5vw, 72px);
  color: rgba(255,255,255,0.62);
  background: var(--dark);
  font-size: 0.9rem;
}
@media (max-width: 920px) {
  .shell { display: block; width: min(100% - 20px, 760px); padding-top: 10px; }
  .sidebar { position: relative; top: auto; max-height: 320px; margin-bottom: 14px; }
  .insights, .topic-grid { grid-template-columns: 1fr; }
  .topbar { position: relative; align-items: flex-start; flex-direction: column; }
  .hero { min-height: 76vh; }
  h1 { font-size: clamp(2.8rem, 16vw, 5.6rem); }
}
"""


def js(search_index: list[dict[str, str]]) -> str:
    return "const SEARCH_INDEX = " + json.dumps(search_index, ensure_ascii=False) + r""";
const input = document.querySelector('[data-search]');
const results = document.querySelector('[data-results]');
function renderResults(query) {
  const q = query.trim().toLowerCase();
  if (!q) {
    results.style.display = 'none';
    results.innerHTML = '';
    return;
  }
  const hits = SEARCH_INDEX
    .map(item => ({ item, score: item.text.toLowerCase().indexOf(q) }))
    .filter(hit => hit.score !== -1)
    .slice(0, 8);
  results.innerHTML = hits.length
    ? hits.map(hit => `<a href="#${hit.item.anchor}">${hit.item.title}</a>`).join('')
    : '<p>No matches found.</p>';
  results.style.display = 'block';
}
input?.addEventListener('input', event => renderResults(event.target.value));
const progress = document.querySelector('.progress');
function updateProgress() {
  const max = document.documentElement.scrollHeight - window.innerHeight;
  const pct = max > 0 ? (window.scrollY / max) * 100 : 0;
  progress.style.width = `${pct}%`;
}
document.addEventListener('scroll', updateProgress, { passive: true });
updateProgress();
"""


def html_document(title: str, source_name: str, blocks: list[Block], hero_image: str = "") -> str:
    toc = build_toc(blocks)
    sections = build_sections(blocks)
    word_count = sum(len(re.findall(r"\w+", block.text, flags=re.UNICODE)) for block in blocks)
    lang = page_language(title, blocks)
    copy = ui_copy(lang)
    hero_image = hero_image or first_image_src(blocks)
    hero_style = f' style="--hero-image: url(\'{html.escape(hero_image)}\')"' if hero_image else ""
    description_source = next((section["text"] for section in sections if section["text"].strip()), "")
    description = compact_text(f"{title}. {description_source}", 155)
    lead = compact_text(description_source, 240) or copy["fallback_lead"]
    toc_html = "\n".join(
        f'<a class="level-{item["level"]}" href="#{html.escape(str(item["anchor"]))}">{html.escape(str(item["title"]))}</a>'
        for item in toc
    )
    body_html = "\n".join(block.html for block in blocks)
    cards_html = "\n".join(
        f"""<article class="insight">
          <div class="kicker">{html.escape(card["kicker"])}</div>
          <h3>{html.escape(card["title"])}</h3>
          <p>{html.escape(card["text"])}</p>
        </article>"""
        for card in insight_cards(sections, toc, word_count, lang)
    )
    topics_html = "\n".join(
        f"""<article class="topic-card">
          <div>
            <div class="kicker">{html.escape(copy["chapter"])} {index:02d}</div>
            <h3>{html.escape(section["title"])}</h3>
            <p>{html.escape(compact_text(section["text"], 170))}</p>
          </div>
          <a href="#{html.escape(section["anchor"])}">{html.escape(copy["read_section"])}</a>
        </article>"""
        for index, section in enumerate(top_sections(sections), 1)
    )
    search_index = [
        {
            "title": section["title"],
            "anchor": section["anchor"],
            "text": f'{section["title"]} {section["text"]}',
        }
        for section in sections
    ]
    search_label = html.escape(copy["search"])
    toc_label = html.escape(copy["toc"])
    og_image = f'\n  <meta property="og:image" content="{html.escape(hero_image)}">' if hero_image else ""
    return f"""<!doctype html>
<html lang="{html.escape(lang)}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <meta name="description" content="{html.escape(description)}">
  <meta property="og:title" content="{html.escape(title)}">
  <meta property="og:description" content="{html.escape(description)}">
  <meta property="og:type" content="website">
  {og_image}
  <meta name="twitter:card" content="summary_large_image">
  <style>{css()}</style>
</head>
<body>
  <div class="progress" aria-hidden="true"></div>
  <header class="topbar">
    <div class="brand">Doc2Web</div>
    <nav class="toplinks" aria-label="Primary navigation">
      <a href="#insights">{html.escape(copy["nav_insights"])}</a>
      <a href="#themes">{html.escape(copy["nav_themes"])}</a>
      <a href="#document">{html.escape(copy["nav_document"])}</a>
    </nav>
  </header>
  <section class="hero" id="top"{hero_style}>
    <div class="hero-inner">
      <div class="eyebrow">{html.escape(copy["eyebrow"])}</div>
      <h1>{html.escape(title)}</h1>
      <p class="dek">{html.escape(lead)}</p>
      <div class="meta">
        <span>{html.escape(source_name)}</span>
        <span>{len(toc)} {html.escape(copy["sections"])}</span>
        <span>{word_count:,} {html.escape(copy["words"])}</span>
        <span>{reading_time(word_count)} {html.escape(copy["read_minutes"])}</span>
      </div>
    </div>
  </section>
  <section class="band" id="insights">
    <div class="section-head">
      <div class="label">{html.escape(copy["insights_label"])}</div>
      <h2>{html.escape(copy["insights_title"])}</h2>
      <p>{html.escape(copy["insights_text"])}</p>
    </div>
    <div class="insights">{cards_html}</div>
  </section>
  <section class="band dark" id="themes">
    <div class="section-head">
      <div class="label">{html.escape(copy["themes_label"])}</div>
      <h2>{html.escape(copy["themes_title"])}</h2>
      <p>{html.escape(copy["themes_text"])}</p>
    </div>
    <div class="topic-grid">{topics_html}</div>
  </section>
  <section class="band" id="document">
    <div class="section-head">
      <div class="label">{html.escape(copy["document_label"])}</div>
      <h2>{html.escape(copy["document_title"])}</h2>
      <p>{html.escape(copy["document_text"])}</p>
    </div>
    <div class="shell">
    <aside class="sidebar">
      <input class="search" data-search type="search" placeholder="{search_label}" aria-label="{search_label}">
      <div class="results" data-results></div>
      <nav class="toc" aria-label="{toc_label}">{toc_html}</nav>
    </aside>
    <main>
      <article class="content">{body_html}</article>
    </main>
  </div>
  </section>
  <footer class="footer">{html.escape(copy["footer_prefix"])} {html.escape(source_name)} {html.escape(copy["footer_suffix"])}</footer>
  <script>{js(search_index)}</script>
</body>
</html>
"""


def convert_doc_to_docx(path: Path, temp_dir: Path) -> Path:
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise SystemExit("Old .doc files require LibreOffice/soffice. Please export the file as .docx and rerun.")
    cmd = [binary, "--headless", "--convert-to", "docx", "--outdir", str(temp_dir), str(path)]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    converted = temp_dir / f"{path.stem}.docx"
    if not converted.exists():
        candidates = list(temp_dir.glob("*.docx"))
        if candidates:
            return candidates[0]
        raise SystemExit("LibreOffice did not produce a .docx file.")
    return converted


def load_blocks(docx_path: Path, out_dir: Path) -> list[Block]:
    with zipfile.ZipFile(docx_path) as zipf:
        rels = parse_rels(zipf, "word/_rels/document.xml.rels")
        ctx = RenderContext(zipf=zipf, out_dir=out_dir, rels=rels, styles=parse_styles(zipf))
        used: set[str] = {"top"}
        blocks = blocks_from_part(ctx, "word/document.xml", used)
        blocks.extend(note_blocks(ctx, "word/footnotes.xml", "Footnotes", used))
        blocks.extend(note_blocks(ctx, "word/endnotes.xml", "Endnotes", used))
        blocks.extend(header_footer_blocks(ctx, used))
        return blocks


def clean_output(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in ("index.html", "assets"):
        target = out_dir / name
        if target.is_dir():
            shutil.rmtree(target)
        elif target.exists():
            target.unlink()
    (out_dir / "assets").mkdir(parents=True, exist_ok=True)


def infer_title(path: Path, blocks: list[Block], explicit: str | None) -> str:
    if explicit:
        return explicit
    for block in blocks:
        if block.kind == "heading" and block.level == 1 and block.title:
            return block.title
    for block in blocks:
        if block.text:
            words = block.text[:90].strip()
            return words if len(words) >= 8 else path.stem
    return path.stem


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert a Word document into a polished static website.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(
            """\
            Examples:
              python3 scripts/doc2web.py report.docx --out site
              python3 scripts/doc2web.py report.doc --out site --title "Annual Report"
            """
        ),
    )
    parser.add_argument("document", type=Path, help="Input .docx or .doc file")
    parser.add_argument("--out", type=Path, default=Path("doc2web-site"), help="Output site directory")
    parser.add_argument("--title", help="Override the page title")
    parser.add_argument("--hero-image", type=Path, help="Use a generated premium hero image instead of falling back to document images")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    source = args.document.expanduser().resolve()
    if not source.exists():
        raise SystemExit(f"Input file not found: {source}")
    clean_output(args.out)
    with tempfile.TemporaryDirectory(prefix="doc2web-") as tmp:
        docx_path = source
        if source.suffix.lower() == ".doc":
            docx_path = convert_doc_to_docx(source, Path(tmp))
        elif source.suffix.lower() != ".docx":
            raise SystemExit("Input must be a .docx file, or a .doc file when LibreOffice is available.")
        blocks = load_blocks(docx_path, args.out)
    title = infer_title(source, blocks, args.title)
    hero_image = install_hero_image(args.hero_image, args.out)
    (args.out / "index.html").write_text(html_document(title, source.name, blocks, hero_image), encoding="utf-8")
    print(f"Created {args.out / 'index.html'}")
    print(f"Assets: {args.out / 'assets'}")
    print(f"Blocks: {len(blocks)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
