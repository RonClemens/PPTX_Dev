from .ooxml import PptxPackage
from .parser import build_document_model
from .markdown_render import (
    render_markdown,
    build_chunks,
    context_markdown_for_comment,
    comment_locations,
    find_comment_anchor,
)
from .anchor import locate_quote_in_paragraph, locate_quote_in_shape, locate_quote_in_slide
from . import writer

__all__ = [
    "PptxPackage",
    "build_document_model",
    "render_markdown",
    "build_chunks",
    "context_markdown_for_comment",
    "comment_locations",
    "find_comment_anchor",
    "locate_quote_in_paragraph",
    "locate_quote_in_shape",
    "locate_quote_in_slide",
    "writer",
]
