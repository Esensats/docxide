"""
docx_studio — Hybrid Declarative VDOM & Fluent Builder for Diploma Generation
Based on DOCX_FRAMEWORK_SPEC_PROPOSAL_v3.md
"""

from .styles import StyleToken, StyleSheet, gost_academic_theme
from .elements import (
    Element, Document, Section, Heading, Paragraph, Span,
    Table, Row, Cell, Image, PlantUml, CodeBlock, PageBreak, EmptyLine, Field
)
from .engine import LayoutEngine, RenderContext
from .builder import DocumentBuilder

__all__ = [
    'StyleToken', 'StyleSheet', 'gost_academic_theme',
    'Element', 'Document', 'Section', 'Heading', 'Paragraph', 'Span',
    'Table', 'Row', 'Cell', 'Image', 'PlantUml', 'CodeBlock', 'PageBreak', 'EmptyLine', 'Field',
    'LayoutEngine', 'RenderContext',
    'DocumentBuilder',
]
