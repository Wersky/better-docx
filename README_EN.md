# better-docx

English | [中文](README.md)

> A high-level Word document automation library built on [python-docx](https://github.com/python-openxml/python-docx) — 42 functions covering tables, images, layout, sections, TOC fields, parsing, and navigation, producing **real native Word objects** (editable `<w:tbl>` tables, embedded pictures), not text-symbol mockups.

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.0-brightgreen)](pyproject.toml)

---

## Why This Library

`python-docx` is a low-level library: many everyday tasks require hand-written OOXML XML, and there are plenty of pitfalls (font fallback, black cell shading, image part errors, runaway page breaks...). (After all, nobody wants to see a mysterious table made of '-' and '|' characters in their generated documents...)

better-docx wraps all of those hard-won lessons into **ready-to-use functions**:

| Task | python-docx raw | better-docx |
|---|---|---|
| CJK fonts | Must manually set `w:eastAsia` or Chinese falls back to the default western font | `set_font(run, "宋体")` sets both font slots |
| Cell shading | Hand-write `w:shd`; wrong `val` renders the cell black | `shade_cell(cell, "D9E2F3")` |
| Dynamic page number | Hand-write the begin/instrText/end three-part field | `add_page_number_field(footer.paragraphs[0])` |
| Auto table of contents | Not supported out of the box | `add_toc_field(doc)` inserts a TOC field; press F9 in Word |
| Mixed page orientation | Sections + manual width/height swap + hidden traps | `set_orientation(sec, landscape=True)` idempotent |
| Inserting images | `add_picture` often fails with missing-`part` errors | `add_image(doc, "x.png")` |
| Locating a paragraph | Linear scan only | `find_heading(doc, "第三章 结论")` anchors by heading |

## Features

- **Native tables**: border styles, cell shading, merged cells, fixed column widths, repeating header rows across pages, no-split rows
- **Image insertion**: body / table-cell placement, centered captions, size control, safe replacement of existing images
- **CJK typography**: one-call font setup for both western and `w:eastAsia` slots, 2-character first-line indent (character units, scales with font size), line spacing and paragraph spacing
- **Page layout**: multi-section portrait/landscape mixing, per-section headers & footers, different first / odd-even pages, live `PAGE` / `NUMPAGES` fields, A4 page setup
- **Paragraphs & lists**: bullets, numbers, nested levels (built-in style system)
- **Style system**: Heading 1–9 outline levels, global style overrides, custom styles, TOC field auto-directory
- **Navigation & perception**: docx → JSON tree / Markdown, anchor-based location, targeted insertion of paragraphs & tables, bookmark and hyperlink read/write
- **Troubleshooting**: diagnose and repair pagination anomalies ("one paragraph per page")

## Installation

```bash
pip install better-docx
```

Or install from source in editable mode:

```bash
git clone https://github.com/Wersky/better-docx.git
cd better-docx
pip install -e .
```

The only dependency is `python-docx>=1.1.0`.

## Quick Start

```python
from docx import Document
from better_docx import set_font, shade_cell, set_cell_text, add_toc_field, add_page_number_field, set_orientation

doc = Document()
sec = doc.sections[0]
set_orientation(sec, landscape=False)                 # A4 portrait (auto swaps w/h, idempotent)
add_page_number_field(sec.footer.paragraphs[0])       # live footer page number "第 X 页"
add_toc_field(doc)                                    # TOC field (use Heading styles, press F9 in Word)

doc.add_heading("一、概述", level=1)                   # outline heading picked up by the TOC
t = doc.add_table(rows=2, cols=2)                     # native table
t.style = "Table Grid"
set_cell_text(t.rows[0].cells[0], "表头", bold=True)   # centered horizontally & vertically
shade_cell(t.rows[0].cells[0], "D9E2F3")              # light-blue header shading

doc.save("report.docx")
```

## API Overview (42 Functions, Three Modules)

`from better_docx import *` re-exports every function; you can also import per-module.

### `better_docx.helpers` — Basic Elements

| Function | Purpose |
|---|---|
| `set_font(run, name, size, bold, italic, color)` | Set run font (auto-sets the `w:eastAsia` slot for CJK) |
| `add_para(doc, text, ...)` / `add_heading(doc, text, level)` | Paragraph / heading (simplified Chinese 黑体 style) |
| `shade_cell(cell, hex)` | Cell background shading |
| `set_cell_text(cell, text, ...)` / `add_para_to_cell(cell, ...)` | Write cell text / append a paragraph in a cell |
| `new_image_paragraph(doc)` / `add_image(doc, path, ...)` / `add_image_to_cell(cell, ...)` | Insert images (body / table cell) |
| `remove_images_from_paragraph(p)` / `global_replace(doc, old, new)` | Replace images / global format-preserving text replace |
| `set_cell_widths(table, widths)` / `set_repeat_header(table)` / `set_row_no_split(row)` | Column widths / repeating header / no-split row |
| `diagnose_pagination(doc)` / `fix_one_paragraph_per_page(doc)` | Detect / repair pagination anomalies |

### `better_docx.layout` — Pages & Typography

| Function | Purpose |
|---|---|
| `set_page_a4(section, margins)` | A4 paper size + margins |
| `set_orientation(section, landscape)` | Portrait/landscape switch (idempotent, auto width/height swap) |
| `enable_even_odd_headers(section)` | Enable different odd/even pages (XML-level) |
| `add_page_number_field(paragraph, ...)` / `add_numpages_field(paragraph, ...)` | Live `PAGE` / `NUMPAGES` fields |
| `set_first_line_indent_chars(paragraph, chars)` | CJK first-line indent of N characters |
| `set_line_spacing(paragraph, value, rule)` | Line spacing (single / 1.5 / exact) |
| `add_bullets(doc, items)` / `add_numbered(doc, pairs)` | Bulleted / numbered lists (nested) |
| `restyle_heading(style, ...)` | Override built-in Heading styles |
| `add_toc_field(doc, levels)` | Insert a TOC field (generate with F9 in Word) |

### `better_docx.navigator` — Parsing & Location

| Function | Purpose |
|---|---|
| `docx_to_json(doc)` / `docx_to_markdown(doc)` | Structured full-document parsing |
| `build_heading_index(doc)` / `find_heading(doc, text)` / `find_paragraph_containing(doc, sub)` | Anchor-based location |
| `insert_paragraph_after(p, ...)` / `insert_paragraph_before_cjk(p, ...)` | Targeted paragraph insertion |
| `insert_table_before(p, ...)` / `insert_table_after(p, ...)` | Targeted table insertion |
| `add_bookmark(p, name)` / `find_bookmark_paragraph(doc, name)` / `append_text_to_bookmark(doc, name, text)` | Bookmark read/write |
| `add_hyperlink(p, url, text)` / `list_hyperlinks(doc)` / `list_bookmarks(doc)` | Hyperlink / bookmark enumeration |

## Complete Example

A multi-section research report (cover → TOC → body → landscape appendix → portrait appendix), see [`examples/make_report.py`](examples/make_report.py):

```python
from docx import Document
from docx.enum.section import WD_SECTION
from better_docx import *

doc = Document()
add_toc_field(doc)                      # TOC field

sec1 = doc.sections[0]                  # portrait section: cover (no header on first page) + body
set_page_a4(sec1)
sec1.different_first_page_header_footer = True
enable_even_odd_headers(sec1)
sec1.header.is_linked_to_previous = False
sec1.header.paragraphs[0].text = "2026 研究报告"
add_page_number_field(sec1.footer.paragraphs[0]); add_numpages_field(sec1.footer.paragraphs[0])

sec2 = doc.add_section(WD_SECTION.NEW_PAGE)   # landscape section for a wide table
set_page_a4(sec2); set_orientation(sec2, landscape=True)
sec2.header.is_linked_to_previous = False
sec2.header.paragraphs[0].text = "附录 A · 横向数据表"

# ... headings, indents, lists, tables, images, bookmarks, hyperlinks

# Agent-style loop: locate with navigator and modify precisely
docx = Document("report.docx")
assert find_heading(docx, "三、结论") is not None
insert_table_before(find_heading(docx, "三、结论"), rows=2, cols=3)
print(docx_to_markdown(docx)[:200])     # parse back to Markdown for preview
```

## Testing

Tests cover all 42 functions (including edge cases and adversarial inputs), run with pytest:

```bash
pip install -e ".[dev]"     # installs pytest
pytest tests/ -v            # or: python -m pytest tests/ -v
```

## Design Notes

- **Real native objects**: tables are `<w:tbl>`, images are embedded drawings, page numbers are field codes (`w:fldChar`) — everything stays editable, updatable and printable in Word. No markdown-symbol mockups.
- **CJK made easy**: every font-related function sets both the western slot and `w:eastAsia`, avoiding Chinese font fallback entirely.
- **Idempotent & robust**: orientation switching, odd/even enablement and more are idempotent; missing images, unusual style names and other edge cases degrade safely instead of raising.
- **Machine-readable**: built-in `docx_to_json` / `docx_to_markdown` make it easy to "perceive → locate → modify → re-check" in agent workflows.

## FAQ

**How does this relate to python-docx?** python-docx is the only dependency — better-docx is a high-level wrapper on top of it, not a fork. You can drop back to the raw API at any time.

**Does it support Chinese?** Chinese is a first-class citizen: SimSun/SimHei/FangSong dual font slots, character-unit indents, and Chinese page-number formats (`CHINESENUM1`).

**Can it read existing documents?** Yes. `docx_to_json` / `docx_to_markdown` parse any .docx; `find_heading` / `find_paragraph_containing` locate content. Note: legacy `.doc` (binary) files must be converted to `.docx` first.

## License

[MIT](LICENSE) © 2026 Wersky