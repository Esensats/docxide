"""Immutable style tokens and theme system."""

from dataclasses import dataclass, replace
from typing import Any, Optional, Dict
from docx.enum.text import WD_ALIGN_PARAGRAPH


@dataclass(frozen=True)
class StyleToken:
    """Immutable design token. No mutations; replace() creates new instances."""
    font_name: str = "Times New Roman"
    font_size: float = 14.0            # pt
    line_spacing: float = 1.0          # single
    bold: bool = False
    italic: bool = False
    color: str = "000000"              # hex RGB
    alignment: Any = WD_ALIGN_PARAGRAPH.JUSTIFY
    indent_first_line: Optional[float] = 1.25  # cm, None = no indent
    indent_left: float = 0.0           # cm
    indent_right: float = 0.0          # cm
    space_before: float = 0.0          # pt
    space_after: float = 0.0           # pt
    page_break_before: bool = False

    def merge(self, **overrides) -> 'StyleToken':
        """Return a new token with overrides applied."""
        return replace(self, **overrides)


class StyleSheet:
    """Theme that maps element roles to base StyleToken."""
    def __init__(self):
        self._rules: Dict[str, StyleToken] = {}
        self._page_width_cm: float = 21.0
        self._page_height_cm: float = 29.7
        self._margin_top_cm: float = 2.0
        self._margin_bottom_cm: float = 2.5
        self._margin_left_cm: float = 3.0
        self._margin_right_cm: float = 1.0

    def register(self, role: str, token: StyleToken):
        self._rules[role] = token

    def get(self, role: str) -> StyleToken:
        return self._rules.get(role, StyleToken())

    @property
    def page_width_cm(self) -> float:
        return self._page_width_cm

    @property
    def page_height_cm(self) -> float:
        return self._page_height_cm

    @property
    def content_width_cm(self) -> float:
        return self._page_width_cm - self._margin_left_cm - self._margin_right_cm

    @property
    def margins(self) -> dict:
        return {
            'top': self._margin_top_cm,
            'bottom': self._margin_bottom_cm,
            'left': self._margin_left_cm,
            'right': self._margin_right_cm,
        }


def gost_academic_theme() -> StyleSheet:
    """GOST-style academic theme for diploma documents."""
    sheet = StyleSheet()

    # Base body text
    sheet.register("body", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        line_spacing=1.0,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        indent_first_line=1.25,
        space_before=0,
        space_after=0,
    ))

    # Title page elements
    sheet.register("title_center", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first_line=None,
    ))

    sheet.register("title_normal", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first_line=None,
    ))

    # Headings
    sheet.register("heading_1", StyleToken(
        font_name="Times New Roman",
        font_size=16.0,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first_line=None,
        space_before=24,
        space_after=12,
        page_break_before=True,
    ))

    sheet.register("heading_2", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        indent_first_line=1.25,
        space_before=12,
        space_after=6,
    ))

    sheet.register("heading_3", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        indent_first_line=1.25,
        space_before=6,
        space_after=3,
    ))

    # Table elements
    sheet.register("table_header", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first_line=None,
    ))

    sheet.register("table_cell", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        indent_first_line=None,
    ))

    # Code
    sheet.register("code", StyleToken(
        font_name="Courier New",
        font_size=10.0,
        line_spacing=1.0,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        indent_first_line=None,
    ))

    # Caption (figure/table captions)
    sheet.register("caption", StyleToken(
        font_name="Times New Roman",
        font_size=12.0,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first_line=None,
        space_before=6,
        space_after=12,
    ))

    # TOC
    sheet.register("toc_title", StyleToken(
        font_name="Times New Roman",
        font_size=16.0,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first_line=None,
        space_before=0,
        space_after=12,
    ))

    # List item
    sheet.register("list_item", StyleToken(
        font_name="Times New Roman",
        font_size=14.0,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        indent_first_line=1.25,
    ))

    return sheet
