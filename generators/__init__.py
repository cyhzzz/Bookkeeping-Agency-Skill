"""
报告生成器
同时生成MD存档和Word交付文档
"""

from .md_generator import MarkdownGenerator
from .docx_generator import DocxGenerator

__all__ = ['MarkdownGenerator', 'DocxGenerator']
