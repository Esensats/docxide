"""Fluent Builder API — primary user interface."""

from typing import List, Union, Optional
from .styles import StyleToken, StyleSheet, gost_academic_theme
from .elements import (
    Document, Section, Heading, Paragraph, Span,
    Table, Row, Cell, Image, PlantUml, CodeBlock, PageBreak, EmptyLine, Field
)
from .engine import LayoutEngine


class DocumentBuilder:
    """Fluent API for building document VDOM trees."""
    def __init__(self, style_sheet: StyleSheet = None):
        self.style_sheet = style_sheet or gost_academic_theme()
        self._root = Document()
        self._current_section: Optional[Section] = None

    def add_section(self, margins: dict = None, orientation: str = 'portrait',
                     footer: str = None) -> 'DocumentBuilder':
        """Add a new section with page settings and optional footer."""
        sec = Section(margins=margins, orientation=orientation, footer=footer)
        self._root.children.append(sec)
        self._current_section = sec
        return self

    def add_heading(self, text: str, level: int = 1, **style_overrides) -> 'DocumentBuilder':
        """Add a heading (level 1-3)."""
        self._ensure_section()
        self._current_section.children.append(Heading(text, level, **style_overrides))
        return self

    def add_paragraph(self, *content, style: str = None, **style_overrides) -> 'DocumentBuilder':
        """Add a paragraph with inline content (strings and Spans)."""
        self._ensure_section()
        inline = []
        for item in content:
            if isinstance(item, str):
                inline.append(item)
            elif isinstance(item, Span):
                inline.append(item)
            else:
                inline.append(str(item))
        self._current_section.children.append(Paragraph(inline, style=style, **style_overrides))
        return self

    def add_table(self, headers: List[str] = None, rows: List[List] = None,
                  visible_borders: bool = True, **style_overrides) -> 'DocumentBuilder':
        """Add a table with headers and data rows."""
        self._ensure_section()
        self._current_section.children.append(
            Table(headers=headers, rows=rows, visible_borders=visible_borders, **style_overrides)
        )
        return self

    def add_image(self, path: str, width_cm: float = None, caption: str = None,
                  **style_overrides) -> 'DocumentBuilder':
        """Add an image with optional caption."""
        self._ensure_section()
        self._current_section.children.append(
            Image(path=path, width_cm=width_cm, caption=caption, **style_overrides)
        )
        return self

    def add_plantuml(self, filename: str, width_cm: float = None, caption: str = None,
                     **style_overrides) -> 'DocumentBuilder':
        """Add a PlantUML diagram (compiles .puml → PNG)."""
        self._ensure_section()
        self._current_section.children.append(
            PlantUml(filename=filename, width_cm=width_cm, caption=caption, **style_overrides)
        )
        return self

    def add_code_block(self, code: str, **style_overrides) -> 'DocumentBuilder':
        """Add a monospace code block."""
        self._ensure_section()
        self._current_section.children.append(CodeBlock(code, **style_overrides))
        return self

    def add_page_break(self) -> 'DocumentBuilder':
        """Add an explicit page break."""
        self._ensure_section()
        self._current_section.children.append(PageBreak())
        return self

    def add_empty_line(self, count: int = 1, style: str = None) -> 'DocumentBuilder':
        """Add empty lines for spacing."""
        self._ensure_section()
        self._current_section.children.append(EmptyLine(count, style=style))
        return self

    def add_field(self, field_code: str) -> 'DocumentBuilder':
        """Add a dynamic field (PAGE, TOC, etc.)."""
        self._ensure_section()
        self._current_section.children.append(Field(field_code))
        return self

    def add_toc(self) -> 'DocumentBuilder':
        """Add a Table of Contents field."""
        self._ensure_section()
        # Add TOC title
        self.add_heading("СОДЕРЖАНИЕ", level=1)
        # Add TOC field
        self._current_section.children.append(Field('TOC \\o "1-3"'))
        return self

    def build(self) -> Document:
        """Build and return the VDOM tree."""
        return self._root

    def save(self, filepath: str) -> None:
        """Build and save the document."""
        engine = LayoutEngine(self.style_sheet)
        engine.render(self._root, filepath)

    def _ensure_section(self):
        """Auto-create a default section if none exists."""
        if not self._current_section:
            self.add_section()
