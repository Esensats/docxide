"""VDOM Element tree — passive data holders."""

from abc import ABC, abstractmethod
from typing import List, Union, Dict


class Element(ABC):
    """Base class for all document parts."""
    def __init__(self, **style_overrides):
        self.style_overrides = style_overrides
        self.children: List['Element'] = []

    @abstractmethod
    def render(self, engine, ctx):
        pass


class Document(Element):
    """Root document container."""
    def __init__(self):
        super().__init__()
        self.children: List[Section] = []

    def render(self, engine, ctx):
        for child in self.children:
            child.render(engine, ctx)


class Section(Element):
    """Page section with margins and optional custom footer."""
    def __init__(self, margins: Dict[str, float] = None, orientation: str = 'portrait',
                 footer: str = None):
        super().__init__()
        self.margins = margins or {'top': 2, 'bottom': 2.5, 'left': 3, 'right': 1}
        self.orientation = orientation
        self.footer = footer  # None=default page num, "page_number"=page num, str=custom text

    def render(self, engine, ctx):
        engine.render_section(self, ctx)


class Heading(Element):
    """Heading with auto page break for level 1."""
    def __init__(self, text: str, level: int = 1, **style_overrides):
        super().__init__(**style_overrides)
        self.text = text
        self.level = level

    def render(self, engine, ctx):
        engine.render_heading(self, ctx)


class Paragraph(Element):
    """Paragraph with inline content (strings and Spans)."""
    def __init__(self, content: Union[str, List[Union[str, 'Span']]], style: str = None, **style_overrides):
        super().__init__(**style_overrides)
        if isinstance(content, str):
            self.content: List[Union[str, 'Span']] = [content]
        else:
            self.content = content
        self.style_role = style  # e.g., "title_body", "title_bold"

    def render(self, engine, ctx):
        engine.render_paragraph(self, ctx)


class Span(Element):
    """Inline formatting run."""
    def __init__(self, text: str, **style_overrides):
        super().__init__(**style_overrides)
        self.text = text

    def render(self, engine, ctx):
        pass  # Rendered inside Paragraph


class Table(Element):
    """Table with rows of cells."""
    def __init__(self, headers: List[str] = None, rows: List[List] = None,
                 visible_borders: bool = True, **style_overrides):
        super().__init__(**style_overrides)
        self.headers = headers or []
        self.rows = rows or []
        self.visible_borders = visible_borders

    def render(self, engine, ctx):
        engine.render_table(self, ctx)


class Row(Element):
    """Table row."""
    def __init__(self, cells: List['Cell'] = None):
        super().__init__()
        self.cells = cells or []

    def render(self, engine, ctx):
        pass  # Rendered inside Table


class Cell(Element):
    """Table cell with content."""
    def __init__(self, content: Union[str, 'Span', List[Element]], width_cm: float = None,
                 bold: bool = False, alignment=None, **style_overrides):
        super().__init__(**style_overrides)
        if isinstance(content, str):
            self.content = [Paragraph(content)]
        elif isinstance(content, Span):
            # Wrap Span in a Paragraph
            self.content = [Paragraph([content])]
        elif isinstance(content, list):
            self.content = content
        else:
            self.content = [Paragraph(str(content))]
        self.width_cm = width_cm
        self.bold = bold
        self.alignment = alignment

    def render(self, engine, ctx):
        pass  # Rendered inside Row


class PlantUml(Element):
    """PlantUML diagram rendered from a .puml file."""
    def __init__(self, filename: str, width_cm: float = None, caption: str = None,
                 **style_overrides):
        super().__init__(**style_overrides)
        self.filename = filename
        self.width_cm = width_cm
        self.caption = caption

    def render(self, engine, ctx):
        engine.render_plantuml(self, ctx)


class Image(Element):
    """Block image."""
    def __init__(self, path: str, width_cm: float = None, caption: str = None,
                 **style_overrides):
        super().__init__(**style_overrides)
        self.path = path
        self.width_cm = width_cm
        self.caption = caption

    def render(self, engine, ctx):
        engine.render_image(self, ctx)


class CodeBlock(Element):
    """Monospace code block."""
    def __init__(self, code: str, **style_overrides):
        super().__init__(**style_overrides)
        self.code = code

    def render(self, engine, ctx):
        engine.render_code_block(self, ctx)


class PageBreak(Element):
    """Explicit page break."""
    def __init__(self):
        super().__init__()

    def render(self, engine, ctx):
        engine.render_page_break(ctx)


class EmptyLine(Element):
    """Empty paragraph for spacing."""
    def __init__(self, count: int = 1, style: str = None):
        super().__init__()
        self.count = count
        self.style_role = style

    def render(self, engine, ctx):
        engine.render_empty_line(self, ctx)


class Field(Element):
    """Dynamic field (e.g., PAGE, TOC)."""
    def __init__(self, field_code: str):
        super().__init__()
        self.field_code = field_code  # e.g., "PAGE", "TOC \\o \"1-3\""

    def render(self, engine, ctx):
        engine.render_field(self, ctx)
