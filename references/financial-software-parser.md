# 财务软件解析器说明

## 支持的软件

| 软件 | 格式 | 特征识别 |
|------|------|---------|
| 金蝶KIS | .docx | 文档内包含"金蝶"、"KIS"字样 |
| 用友U8/U9 | .xlsx | 工作表名为"记账凭证"、"明细账" |
| 畅捷通T+/T6 | .xlsx | 工作表名为"凭证库"、"科目余额表" |

## 解析流程

```
财务软件导出文件
      │
      ▼
  文件类型检测
      │
      ▼
  格式解析（根据软件类型）
      │
      ▼
  数据标准化（统一JSON格式）
      │
      ▼
  返回解析结果给调度器
```

## 金蝶KIS解析规则

金蝶导出的Word文档结构：
- 标题区：公司名称、导出日期、期间
- 表格区：凭证数据表格
- 识别特征：文档属性或内容包含"KIS"或"金蝶"

**解析字段映射**：
| Word表格列 | 标准字段 |
|-----------|---------|
| 日期 | voucher_date |
| 凭证号 | voucher_no |
| 摘要 | summary |
| 科目编码 | account_code |
| 科目名称 | account_name |
| 借方金额 | debit_amount |
| 贷方金额 | credit_amount |

## 用友U8/U9解析规则

用友导出的Excel结构：
- 工作表：记账凭证、明细账、科目余额表等
- 识别特征：文件属性或工作表名称

**解析字段映射**：
| Excel列 | 标准字段 |
|---------|---------|
| 制单日期 | voucher_date |
| 凭证字号 | voucher_no |
| 会计科目 | account_name |
| 金额 | amount |

## 畅捷通T+/T6解析规则

畅捷通导出的Excel结构：
- 工作表：凭证库、科目余额表
- 识别特征：工作表名为"凭证库"

**解析字段映射**：
| Excel列 | 标准字段 |
|---------|---------|
| 凭证日期 | voucher_date |
| 凭证编号 | voucher_no |
| 科目编码 | account_code |
| 科目名称 | account_name |
| 借方 | debit_amount |
| 贷方 | credit_amount |

## 解析结果格式

```json
{
  "parser": "kingdee",
  "company_name": "公司名称",
  "period": "2026-04",
  "vouchers": [
    {
      "date": "2026-04-01",
      "no": "记-001",
      "summary": "收到货款",
      "entries": [
        {"account": "银行存款", "debit": 10000, "credit": 0},
        {"account": "应收账款", "debit": 0, "credit": 10000}
      ]
    }
  ],
  "confidence": 0.95
}
```

## 使用方式

当员工上传财务软件导出文件时，自动调用：

```python
from parsers import ParserRegistry

registry = ParserRegistry()
parser = registry.auto_select(file_path)
result = parser.parse(file_path)
```

## 注意事项

1. 解析器自动识别文件类型，无需员工手动选择
2. 如解析失败，提示员工检查导出文件格式是否标准
3. 解析后的数据供后续模块使用，员工无感知
