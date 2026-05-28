# docxide

Declarative VDOM + fluent builder for `.docx` document generation. Built on `python-docx`.

## Installation

```sh
pip install -e .
```

Requires Python >=3.10.

## Quick Start

```python
from docx_studio import DocumentBuilder, Span

doc = DocumentBuilder()

doc.add_section(footer="page_number") \
   .add_heading("Chapter 1", level=1) \
   .add_paragraph(
       "Plain text with ",
       Span("bold inline", bold=True),
       " content."
   ) \
   .add_table(
       headers=["Name", "Value"],
       rows=[["Alpha", "1"], ["Beta", "2"]]
   ) \
   .save("output.docx")
```

## Architecture

Four files, strict layering. Nothing outside `engine.py` touches `python-docx`.

| Layer | File | Purpose |
|-------|------|---------|
| Data | `styles.py` | `StyleToken` (frozen dataclass) + `StyleSheet` (role → token map) |
| Data | `elements.py` | VDOM nodes — passive, no rendering logic |
| API | `builder.py` | Fluent chain that assembles the VDOM tree |
| Rendering | `engine.py` | `LayoutEngine` (visitor) + `RenderContext` (restricted bridge to python-docx) |

## Style System

Styles are immutable dataclasses. Override via `token.merge(**kwargs)` — never mutate.

```python
from docx_studio.styles import StyleToken

custom = StyleToken(font_size=12, bold=True)
modified = custom.merge(indent_first_line=0)  # new instance, original unchanged
```

Built-in theme: `gost_academic_theme()` — GOST-style academic formatting (Times New Roman 14pt, 1.25cm indent, justified).

### Registered Roles

`body`, `title_center`, `title_normal`, `heading_1`, `heading_2`, `heading_3`, `table_header`, `table_cell`, `code`, `caption`, `toc_title`, `list_item`.

## API Reference

### DocumentBuilder

| Method | Description |
|--------|-------------|
| `add_section(margins, footer)` | New page section. `footer=None` → page number, `footer="text"` → custom text. |
| `add_heading(text, level)` | Heading (1–3). Level 1 auto-adds page break. |
| `add_paragraph(*content, style)` | Paragraph with strings and `Span` objects. |
| `add_table(headers, rows, visible_borders)` | Table. Rows can contain strings, `Cell`, or `Span` objects. |
| `add_image(path, width_cm, caption)` | Block image with optional caption. |
| `add_plantuml(filename, width_cm, caption)` | PlantUML diagram (compiles `.puml` → PNG, cached). |
| `add_code_block(code)` | Monospace code block. |
| `add_page_break()` | Explicit page break. |
| `add_empty_line(count, style)` | Empty paragraphs for spacing. |
| `add_field(field_code)` | Dynamic field (e.g., `"PAGE"`, `"TOC \\o \"1-3\""`) |
| `add_toc(title)` | Table of contents field with title. |
| `save(filepath)` | Build VDOM and write `.docx`. |

### Span

Inline formatting: `Span("text", bold=True, italic=True, color="800000")`.

### Cell

Table cell with rich content: `Cell([Paragraph("text"), Span("bold")], bold=True, width_cm=5)`.

## Default Page Layout

- A4 portrait (21 × 29.7 cm)
- Margins: top 2 cm, bottom 2.5 cm, left 3 cm, right 1 cm
- Font: Times New Roman 14pt, single spacing
- Body: justified, 1.25cm first-line indent

## License

Internal project — no public license.
