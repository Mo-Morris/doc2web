---
name: doc2web
description: Convert large Word documents (.docx, and .doc via LibreOffice when available) into polished, complete static websites with responsive layout, section navigation, tables, images, notes, and search. Use when a user provides a Word document and wants all content preserved in a good-looking website.
metadata:
  short-description: Turn large Word docs into polished websites
---

# Doc2Web

Use this skill when the user wants a Word document converted into a complete, readable website rather than a summary.

## Primary Workflow

1. Put the user's `.docx` or `.doc` file somewhere accessible in the workspace.
2. Run the converter:

```bash
python3 scripts/doc2web.py path/to/document.docx --out path/to/site
```

For old `.doc` files, the script tries `soffice`/`libreoffice` to convert to `.docx` first. If neither is installed, ask the user for a `.docx` export.

3. Inspect `path/to/site/index.html` and the generated `assets/` folder.
4. If the document is important or visually complex, open the result in a browser and check:
   - all major headings appear in the navigation
   - long tables scroll cleanly on mobile
   - images render from `assets/media/`
   - footnotes/endnotes appear near the bottom
   - search finds expected phrases

## Design Guidance

- Preserve all content. Do not summarize or omit sections unless the user explicitly asks.
- Favor a polished editorial site: readable typography, strong hierarchy, generous spacing, sticky section navigation, and responsive tables.
- Large documents should remain navigable: use heading anchors, a progress bar, search, and section cards.
- If the generated hierarchy is weak because the source document has poor styles, improve headings manually in the HTML only after confirming the source structure from the generated page.

## Script Notes

- `scripts/doc2web.py` uses Python standard library only for `.docx` parsing.
- Embedded images are extracted from the Word package into the generated site.
- Hyperlinks, headings, paragraphs, tables, headers/footers, footnotes, and endnotes are included when present.
- The output is static and can be hosted by any web server or opened locally.
