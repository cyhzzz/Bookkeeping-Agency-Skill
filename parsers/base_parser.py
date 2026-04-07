"""
财务软件解析器基类
所有解析器必须实现标准接口
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Dict, Any, Optional


@dataclass
class VoucherEntry:
    """凭证分录"""
    account_code: str
    account_name: str
    debit_amount: float = 0.0
    credit_amount: float = 0.0


@dataclass
class ParsedVoucher:
    """解析后的凭证数据"""
    voucher_id: str
    date: str
    voucher_no: str
    summary: str
    entries: List[VoucherEntry] = field(default_factory=list)
    attachments: int = 0
    source_file: str = ""


@dataclass
class ParsedTrialBalance:
    """解析后的科目余额表"""
    account_code: str
    account_name: str
    opening_balance: float = 0.0
    debit_amount: float = 0.0
    credit_amount: float = 0.0
    closing_balance: float = 0.0


@dataclass
class ParsedFinancialData:
    """解析后的财务数据"""
    parser_name: str
    company_name: str
    period: str
    vouchers: List[ParsedVoucher] = field(default_factory=list)
    trial_balance: List[ParsedTrialBalance] = field(default_factory=list)
    raw_data: Dict[str, Any] = field(default_factory=dict)
    parse_time: str = field(default_factory=lambda: datetime.now().isoformat())
    confidence: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            'parser': self.parser_name,
            'company_name': self.company_name,
            'period': self.period,
            'vouchers': [
                {
                    'id': v.voucher_id,
                    'date': v.date,
                    'no': v.voucher_no,
                    'summary': v.summary,
                    'entries': [
                        {
                            'account_code': e.account_code,
                            'account_name': e.account_name,
                            'debit': e.debit_amount,
                            'credit': e.credit_amount
                        }
                        for e in v.entries
                    ]
                }
                for v in self.vouchers
            ],
            'confidence': self.confidence
        }


class BaseParser(ABC):
    """解析器基类"""

    @property
    @abstractmethod
    def software_name(self) -> str:
        """软件名称"""
        pass

    @property
    @abstractmethod
    def supported_formats(self) -> List[str]:
        """支持的格式"""
        pass

    @abstractmethod
    def can_parse(self, file_path: str) -> bool:
        """判断是否能解析此文件"""
        pass

    @abstractmethod
    def parse(self, file_path: str) -> ParsedFinancialData:
        """解析文件"""
        pass
