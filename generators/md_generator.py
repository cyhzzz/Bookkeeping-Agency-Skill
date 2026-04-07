"""
Markdown报告生成器
用于生成存档用的完整报告
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional
import json


class MarkdownGenerator:
    """Markdown报告生成器"""

    def __init__(self, template_dir: Optional[str] = None):
        self.template_dir = template_dir

    def generate(self, report_type: str, data: Dict[str, Any],
                 output_path: str, template: Optional[str] = None) -> str:
        """
        生成Markdown报告

        Args:
            report_type: 报告类型
            data: 报告数据
            output_path: 输出路径
            template: 可选模板名称

        Returns:
            生成的Markdown文件路径
        """
        # 1. 生成元信息
        metadata = self._generate_metadata(report_type, data)

        # 2. 生成报告内容
        content = self._render_content(report_type, data)

        # 3. 组合完整内容
        full_content = f"{metadata}\n\n{content}"

        # 4. 保存文件
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_content)

        return output_path

    def _generate_metadata(self, report_type: str, data: Dict[str, Any]) -> str:
        """生成报告元信息"""
        return f"""---
report_type: {report_type}
customer_id: {data.get('customer_id', '')}
customer_name: {data.get('customer_name', '')}
period: {data.get('period', '')}
generated_at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
generated_by: AI代账工具
version: 1.0
---"""

    def _render_content(self, report_type: str, data: Dict[str, Any]) -> str:
        """根据报告类型渲染内容"""
        content_func = getattr(self, f'_render_{report_type}', None)
        if content_func:
            return content_func(data)
        return self._render_generic(data)

    def _render_monthly_report(self, data: Dict[str, Any]) -> str:
        """月度经营分析报告"""
        return f"""# {data.get('period', '')}月度经营分析报告

## 一、经营概况

### 1.1 主要财务指标

| 指标 | 本月 | 本年累计 | 环比 | 同比 |
|------|------|----------|------|------|
| 营业收入 | {data.get('revenue', 'N/A')} | {data.get('revenue_ytd', 'N/A')} | {data.get('revenue_mom', 'N/A')} | {data.get('revenue_yoy', 'N/A')} |
| 营业成本 | {data.get('cost', 'N/A')} | {data.get('cost_ytd', 'N/A')} | {data.get('cost_mom', 'N/A')} | {data.get('cost_yoy', 'N/A')} |
| 毛利润 | {data.get('gross_profit', 'N/A')} | {data.get('gross_profit_ytd', 'N/A')} | {data.get('gp_mom', 'N/A')} | {data.get('gp_yoy', 'N/A')} |
| 毛利率 | {data.get('gross_margin', 'N/A')} | {data.get('gross_margin_ytd', 'N/A')} | - | - |
| 净利润 | {data.get('net_profit', 'N/A')} | {data.get('net_profit_ytd', 'N/A')} | {data.get('np_mom', 'N/A')} | {data.get('np_yoy', 'N/A')} |

## 二、税负分析

### 2.1 各税种税负情况

| 税种 | 本月税额 | 税负率 | 行业均值 | 偏差 |
|------|----------|--------|----------|------|
| 增值税 | {data.get('vat_tax', 'N/A')} | {data.get('vat_rate', 'N/A')} | {data.get('vat_benchmark', 'N/A')} | {data.get('vat_deviation', 'N/A')} |
| 企业所得税 | {data.get('eit_tax', 'N/A')} | {data.get('eit_rate', 'N/A')} | {data.get('eit_benchmark', 'N/A')} | {data.get('eit_deviation', 'N/A')} |

## 三、风险提示

{data.get('risk_alerts', '无明显异常')}

## 四、下月注意事项

{data.get('next_month_notes', '请关注申报截止日期')}

---
**报告生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**编制人**：AI代账工具
"""

    def _render_client_onboarding(self, data: Dict[str, Any]) -> str:
        """客户接入档案"""
        return f"""# 客户接入档案

## 一、客户基本信息

| 项目 | 内容 |
|------|------|
| 公司名称 | {data.get('company_name', '')} |
| 统一社会信用代码 | {data.get('unified_credit_code', '')} |
| 法人代表 | {data.get('legal_person', '')} |
| 注册资本 | {data.get('registered_capital', '')} |
| 经营范围 | {data.get('business_scope', '')} |
| 注册地址 | {data.get('registered_address', '')} |

## 二、纳税人类型

- **类型**：{data.get('taxpayer_type', '')}
- **行业分类**：{data.get('industry_category', '')}

## 三、会计科目表

{data.get('accounting_chart', '详见附件')}

## 四、税种核定

| 税种 | 核定情况 |
|------|----------|
| 增值税 | {data.get('vat_status', '')} |
| 企业所得税 | {data.get('eit_status', '')} |
| 个人所得税 | {data.get('it_status', '')} |
| 附加税 | {data.get('附加_tax_status', '')} |

## 五、关键日期

| 日期类型 | 公历日期 | 备注 |
|----------|----------|------|
| 月度申报截止 | 每月15日 | 如遇节假日顺延 |
| 季度所得税申报 | 每季度首月15日 | |
| 年度汇算清缴 | 5月31日 | |

---
**建档时间**：{datetime.now().strftime('%Y-%m-%d')}
**建档人**：AI代账工具
"""

    def _render_risk_alert(self, data: Dict[str, Any]) -> str:
        """风险预警报告"""
        return f"""# 税务风险预警报告

## 一、风险概览

| 指标 | 结果 |
|------|------|
| 综合风险评分 | {data.get('overall_score', '')} |
| 风险等级 | {data.get('risk_level', '')} |
| 评估周期 | {data.get('period', '')} |

## 二、风险明细

{data.get('risk_details', '无明显风险')}

## 三、处理建议

{data.get('recommendations', '持续监控')}

---
**报告生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**评估人**：AI代账工具
"""

    def _render_generic(self, data: Dict[str, Any]) -> str:
        """通用报告模板"""
        return f"""# {data.get('title', '报告')}

## 基本信息

- 客户名称：{data.get('customer_name', '')}
- 报告周期：{data.get('period', '')}
- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 报告内容

{json.dumps(data, ensure_ascii=False, indent=2)}

---
**编制人**：AI代账工具
"""
