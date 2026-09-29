"""Convert DOCX teaching materials to Git-friendly Markdown."""

from .converter import ConversionError, ConversionResult, convert_docx

__all__ = ["ConversionError", "ConversionResult", "convert_docx"]
