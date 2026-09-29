"""Convert DOCX and PDF teaching materials to Git-friendly Markdown."""

from .converter import ConversionError, ConversionResult, convert_document, convert_docx

__all__ = ["ConversionError", "ConversionResult", "convert_docx", "convert_document"]
