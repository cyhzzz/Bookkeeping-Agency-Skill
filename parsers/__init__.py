"""
财务软件解析器注册表
自动识别并调度对应解析器
"""

from pathlib import Path
from typing import List, Optional
from .base_parser import BaseParser, ParsedFinancialData
from .kingdee_parser import KingdeeParser
from .yonyou_parser import YonyouParser
from .chanjet_parser import ChanjetParser


class ParserRegistry:
    """解析器注册表"""

    def __init__(self):
        self.parsers: List[BaseParser] = [
            KingdeeParser(),
            YonyouParser(),
            ChanjetParser()
        ]

    def auto_select(self, file_path: str) -> BaseParser:
        """自动选择合适的解析器"""
        path = Path(file_path)

        for parser in self.parsers:
            if parser.can_parse(file_path):
                return parser

        raise ValueError(
            f"无法解析文件: {file_path}\n"
            f"支持的格式: .docx (金蝶), .xlsx (用友/畅捷通)"
        )

    def parse_financial_file(self, file_path: str) -> ParsedFinancialData:
        """解析财务文件"""
        parser = self.auto_select(file_path)
        return parser.parse(file_path)


__all__ = [
    'ParserRegistry',
    'BaseParser',
    'ParsedFinancialData',
    'KingdeeParser',
    'YonyouParser',
    'ChanjetParser'
]
