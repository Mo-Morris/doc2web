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


def css() -> str:
    return r"""
:root {
  --ink: #17211b;
  --muted: #637067;
  --paper: #fffdf5;
  --panel: rgba(255, 255, 255, 0.76);
  --line: rgba(23, 33, 27, 0.14);
  --accent: #b45f2a;
  --accent-2: #1c6f62;
  --shadow: 0 24px 80px rgba(38, 31, 20, 0.14);
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
  margin: 0;
  color: var(--ink);
  background:
    radial-gradient(circle at top left, rgba(180, 95, 42, 0.20), transparent 34rem),
    radial-gradient(circle at 80% 10%, rgba(28, 111, 98, 0.18), transparent 30rem),
    linear-gradient(135deg, #f7efe0 0%, #fffdf5 44%, #eef5ee 100%);
  font-family: Charter, "Iowan Old Style", "Palatino Linotype", Georgia, serif;
}
.progress {
  position: fixed;
  inset: 0 auto auto 0;
  width: 0;
  height: 4px;
  z-index: 20;
  background: linear-gradient(90deg, var(--accent), var(--accent-2));
}
.shell {
  display: grid;
  grid-template-columns: minmax(230px, 310px) minmax(0, 1fr);
  gap: clamp(1rem, 3vw, 3rem);
  width: min(1480px, calc(100% - 32px));
  margin: 0 auto;
  padding: 32px 0 72px;
}
.sidebar {
  position: sticky;
  top: 24px;
  align-self: start;
  max-height: calc(100vh - 48px);
  overflow: auto;
  padding: 20px;
  border: 1px solid var(--line);
  border-radius: 28px;
  background: rgba(255, 253, 245, 0.72);
  box-shadow: var(--shadow);
  backdrop-filter: blur(18px);
}
.brand {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 18px;
  color: var(--accent);
  font-family: Avenir Next, Futura, Trebuchet MS, sans-serif;
  font-size: 0.75rem;
  font-weight: 800;
  letter-spacing: 0.16em;
  text-transform: uppercase;
}
.brand::before {
  width: 13px;
  height: 13px;
  content: "";
  border-radius: 999px;
  background: var(--accent);
  box-shadow: 18px 0 0 var(--accent-2);
}
.search {
  width: 100%;
  margin-bottom: 16px;
  padding: 12px 14px;
  border: 1px solid var(--line);
  border-radius: 999px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.72);
  font: 600 0.95rem Avenir Next, Futura, Trebuchet MS, sans-serif;
}
.results {
  display: none;
  margin-bottom: 14px;
  padding: 10px;
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.62);
}
.results a, .toc a {
  display: block;
  color: var(--ink);
  text-decoration: none;
}
.results a {
  padding: 8px 10px;
  border-radius: 12px;
  font-family: Avenir Next, Futura, Trebuchet MS, sans-serif;
  font-size: 0.9rem;
}
.results a:hover, .toc a:hover { background: rgba(180, 95, 42, 0.10); }
.toc {
  display: grid;
  gap: 3px;
  padding-top: 10px;
  border-top: 1px solid var(--line);
  font-family: Avenir Next, Futura, Trebuchet MS, sans-serif;
}
.toc a {
  padding: 7px 9px;
  border-radius: 11px;
  color: var(--muted);
  font-size: 0.88rem;
  line-height: 1.3;
}
.toc .level-1 { margin-left: 0; color: var(--ink); font-weight: 800; }
.toc .level-2 { margin-left: 10px; }
.toc .level-3 { margin-left: 20px; font-size: 0.82rem; }
.toc .level-4, .toc .level-5, .toc .level-6 { margin-left: 30px; font-size: 0.78rem; }
.hero {
  margin-bottom: 26px;
  padding: clamp(2rem, 7vw, 5.5rem);
  border: 1px solid var(--line);
  border-radius: 42px;
  background:
    linear-gradient(135deg, rgba(255,255,255,0.80), rgba(255,255,255,0.46)),
    repeating-linear-gradient(135deg, rgba(23,33,27,0.035) 0 1px, transparent 1px 13px);
  box-shadow: var(--shadow);
}
.eyebrow {
  color: var(--accent);
  font: 800 0.78rem Avenir Next, Futura, Trebuchet MS, sans-serif;
  letter-spacing: 0.18em;
  text-transform: uppercase;
}
h1, h2, h3, h4, h5, h6 {
  margin: 1.3em 0 0.46em;
  line-height: 1.08;
  text-wrap: balance;
}
h1 {
  max-width: 14ch;
  margin-top: 0.25em;
  font-size: clamp(3.2rem, 10vw, 8.5rem);
  letter-spacing: -0.08em;
}
h2 { font-size: clamp(2rem, 4vw, 3.8rem); letter-spacing: -0.055em; }
h3 { font-size: clamp(1.55rem, 2.7vw, 2.4rem); letter-spacing: -0.035em; }
h4 { font-size: 1.35rem; }
.meta {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 22px;
  color: var(--muted);
  font: 700 0.86rem Avenir Next, Futura, Trebuchet MS, sans-serif;
}
.meta span {
  padding: 8px 12px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: rgba(255,255,255,0.58);
}
.content {
  padding: clamp(1.2rem, 5vw, 4.5rem);
  border: 1px solid var(--line);
  border-radius: 42px;
  background: var(--panel);
  box-shadow: var(--shadow);
  backdrop-filter: blur(14px);
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
  border-radius: 18px;
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
  border-radius: 28px;
  background: rgba(255,255,255,0.68);
}
.doc-figure img {
  display: block;
  max-width: 100%;
  height: auto;
  margin: 0 auto;
  border-radius: 18px;
}
.table-wrap {
  width: 100%;
  margin: 1.6rem 0;
  overflow-x: auto;
  border: 1px solid var(--line);
  border-radius: 24px;
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
  border-radius: 18px;
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
  font: 800 0.8rem Avenir Next, Futura, Trebuchet MS, sans-serif;
}
@media (max-width: 920px) {
  .shell { display: block; width: min(100% - 20px, 760px); padding-top: 10px; }
  .sidebar { position: relative; top: auto; max-height: 320px; margin-bottom: 14px; border-radius: 24px; }
  .hero, .content { border-radius: 28px; }
  h1 { font-size: clamp(2.8rem, 18vw, 5.6rem); }
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


def html_document(title: str, source_name: str, blocks: list[Block]) -> str:
    toc = build_toc(blocks)
    sections = build_sections(blocks)
    word_count = sum(len(re.findall(r"\w+", block.text, flags=re.UNICODE)) for block in blocks)
    toc_html = "\n".join(
        f'<a class="level-{item["level"]}" href="#{html.escape(str(item["anchor"]))}">{html.escape(str(item["title"]))}</a>'
        for item in toc
    )
    body_html = "\n".join(block.html for block in blocks)
    search_index = [
        {
            "title": section["title"],
            "anchor": section["anchor"],
            "text": f'{section["title"]} {section["text"]}',
        }
        for section in sections
    ]
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>{css()}</style>
</head>
<body>
  <div class="progress" aria-hidden="true"></div>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand">Doc2Web</div>
      <input class="search" data-search type="search" placeholder="Search this document" aria-label="Search this document">
      <div class="results" data-results></div>
      <nav class="toc" aria-label="Table of contents">{toc_html}</nav>
    </aside>
    <main>
      <section class="hero" id="top">
        <div class="eyebrow">Converted Word Document</div>
        <h1>{html.escape(title)}</h1>
        <div class="meta">
          <span>{html.escape(source_name)}</span>
          <span>{len(toc)} headings</span>
          <span>{word_count:,} words</span>
        </div>
      </section>
      <article class="content">{body_html}</article>
    </main>
  </div>
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
    (args.out / "index.html").write_text(html_document(title, source.name, blocks), encoding="utf-8")
    print(f"Created {args.out / 'index.html'}")
    print(f"Assets: {args.out / 'assets'}")
    print(f"Blocks: {len(blocks)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
