"""
出口退税计算脚本

支持两种类型：
- 生产型企业：免抵退计算
- 外贸型企业：先征后退计算

输入：
- 报关单数据（商品编码、出口数量、离岸价、汇率）
- 采购发票数据（商品名称、金额、税率）
- 退税率表

输出：
- 退税金额
- 免抵退汇总表
- 风险预警列表
"""

import json
import argparse
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class CustomsDeclaration:
    """报关单数据"""
    declaration_id: str  # 报关单号
    hs_code: str  # 商品HS编码
    goods_name: str  # 商品名称
    quantity: float  # 出口数量
    fob_usd: float  # 离岸价（美元）
    destination_country: str  # 目的国
    declaration_date: str  # 报关日期


@dataclass
class PurchaseInvoice:
    """采购发票数据"""
    invoice_id: str  # 发票号
    goods_name: str  # 商品名称
    amount_with_tax: float  # 价税合计（含税金额）
    tax_rate: float  # 税率
    invoice_date: str  # 开票日期


@dataclass
class RebateRate:
    """退税率信息"""
    hs_code: str  # 商品HS编码
    rebate_rate: float  # 退税率
    applicable_tax_rate: float  # 适用增值税率


class ExportRebateCalculator:
    """出口退税计算器"""

    # 2024年主要商品退税率参考表（实际以税务局确认为准）
    DEFAULT_REBATE_RATES = {
        "8471": {"rebate_rate": 13, "applicable_tax_rate": 13},  # 计算机
        "8501": {"rebate_rate": 13, "applicable_tax_rate": 13},  # 电机
        "6201": {"rebate_rate": 13, "applicable_tax_rate": 13},  # 服装
        "7204": {"rebate_rate": 0, "applicable_tax_rate": 13},  # 钢铁废料
        "1001": {"rebate_rate": 9, "applicable_tax_rate": 9},   # 小麦
    }

    def __init__(self, enterprise_type: str = "production"):
        """
        初始化计算器

        Args:
            enterprise_type: "production" (生产型) 或 "trading" (外贸型)
        """
        if enterprise_type not in ("production", "trading"):
            raise ValueError("企业类型必须是 'production' (生产型) 或 'trading' (外贸型)")
        self.enterprise_type = enterprise_type
        self.exchange_rate: float = 0.0  # 汇率（申报当期央行中间价）
        self.customs_data: List[CustomsDeclaration] = []
        self.invoice_data: List[PurchaseInvoice] = []
        self.risk_warnings: List[Dict[str, Any]] = []

    def set_exchange_rate(self, rate: float, source: str = "中国人民银行"):
        """
        设置汇率

        Args:
            rate: 汇率（人民币/美元中间价）
            source: 汇率来源
        """
        self.exchange_rate = rate
        self.exchange_rate_source = source

    def add_customs_declaration(self, declaration: CustomsDeclaration):
        """添加报关单数据"""
        self.customs_data.append(declaration)

    def add_purchase_invoice(self, invoice: PurchaseInvoice):
        """添加采购发票数据"""
        self.invoice_data.append(invoice)

    def get_rebate_rate(self, hs_code: str) -> Optional[RebateRate]:
        """
        获取退税率

        Args:
            hs_code: 商品HS编码（前6位）

        Returns:
            退税率信息，未找到返回None
        """
        prefix = hs_code[:6]
        if prefix in self.DEFAULT_REBATE_RATES:
            rates = self.DEFAULT_REBATE_RATES[prefix]
            return RebateRate(
                hs_code=prefix,
                rebate_rate=rates["rebate_rate"],
                applicable_tax_rate=rates["applicable_tax_rate"]
            )
        # 尝试更短的匹配
        for key, rates in self.DEFAULT_REBATE_RATES.items():
            if hs_code.startswith(key):
                return RebateRate(
                    hs_code=key,
                    rebate_rate=rates["rebate_rate"],
                    applicable_tax_rate=rates["applicable_tax_rate"]
                )
        return None

    def calculate_production_type(self) -> Dict[str, Any]:
        """
        生产型企业免抵退计算

        计算公式：
        - 当期免抵退税不得免征和抵扣税额 = 当期出口货物离岸价 × 外汇人民币折合率 ×（出口货物退税率 - 适用税率）
        - 当期应纳税额 = 当期内销货物销项税额 -（当期进项税额 - 当期免抵退税不得免征和抵扣税额）- 上期留抵税额
        - 当期免抵退税额 = 当期出口货物离岸价 × 外汇人民币折合率 × 出口货物退税率
        - 当期应退税额 = min(当期应纳税额绝对值，当期免抵退税额)

        Returns:
            包含各项计算结果的字典
        """
        total_fob_usd = sum(d.fob_usd for d in self.customs_data)
        total_fob_cny = total_fob_usd * self.exchange_rate

        # 按商品汇总，计算不得免征抵扣税额
        total_non_deductible = 0.0
        total_rebate_amount = 0.0
        goods_details = []

        for customs in self.customs_data:
            rebate_info = self.get_rebate_rate(customs.hs_code)
            if rebate_info:
                # 不得免征和抵扣税额
                non_deductible = customs.fob_usd * self.exchange_rate * \
                                 (rebate_info.rebate_rate - rebate_info.applicable_tax_rate) / 100
                # 免抵退税额
                rebate = customs.fob_usd * self.exchange_rate * rebate_info.rebate_rate / 100

                total_non_deductible += non_deductible
                total_rebate_amount += rebate

                goods_details.append({
                    "declaration_id": customs.declaration_id,
                    "goods_name": customs.goods_name,
                    "hs_code": customs.hs_code,
                    "fob_usd": customs.fob_usd,
                    "fob_cny": customs.fob_usd * self.exchange_rate,
                    "rebate_rate": rebate_info.rebate_rate,
                    "applicable_tax_rate": rebate_info.applicable_tax_rate,
                    "non_deductible": non_deductible,
                    "rebate_amount": rebate
                })

        # 计算进项税额（用于计算应纳税额）
        total_input_tax = sum(
            inv.amount_with_tax / (1 + inv.tax_rate / 100) * (inv.tax_rate / 100)
            for inv in self.invoice_data
        )

        # 组装结果
        result = {
            "enterprise_type": "生产型",
            "calculation_details": {
                "total_fob_usd": total_fob_usd,
                "exchange_rate": self.exchange_rate,
                "exchange_rate_source": getattr(self, 'exchange_rate_source', '未知'),
                "total_fob_cny": total_fob_cny,
                "total_input_tax": total_input_tax,
                "total_non_deductible": total_non_deductible,
                "total_rebate_amount": total_rebate_amount
            },
            "goods_details": goods_details,
            "risk_warnings": self.risk_warnings
        }

        return result

    def calculate_trading_type(self) -> Dict[str, Any]:
        """
        外贸型企业先征后退计算

        计算公式：
        - 应退税额 = 收购不含税购进金额 × 退税率
        - 收购不含税购进金额 = 采购发票价税合计 / (1 + 税率)

        Returns:
            包含各项计算结果的字典
        """
        total_rebate = 0.0
        invoice_details = []

        for invoice in self.invoice_data:
            # 查找对应的报关单（按商品名称匹配）
            matched_customs = None
            for customs in self.customs_data:
                if customs.goods_name == invoice.goods_name:
                    matched_customs = customs
                    break

            # 计算不含税购进金额
            purchase_ex_tax = invoice.amount_with_tax / (1 + invoice.tax_rate / 100)

            # 获取退税率
            rebate_rate = 0.0
            hs_code = ""
            if matched_customs:
                rebate_info = self.get_rebate_rate(matched_customs.hs_code)
                if rebate_info:
                    rebate_rate = rebate_info.rebate_rate
                    hs_code = matched_customs.hs_code

            # 计算应退税额
            rebate_amount = purchase_ex_tax * rebate_rate / 100
            total_rebate += rebate_amount

            invoice_details.append({
                "invoice_id": invoice.invoice_id,
                "goods_name": invoice.goods_name,
                "hs_code": hs_code,
                "amount_with_tax": invoice.amount_with_tax,
                "purchase_ex_tax": purchase_ex_tax,
                "tax_rate": invoice.tax_rate,
                "rebate_rate": rebate_rate,
                "rebate_amount": rebate_amount,
                "matched_customs_id": matched_customs.declaration_id if matched_customs else None
            })

        result = {
            "enterprise_type": "外贸型",
            "calculation_details": {
                "total_rebate_amount": total_rebate,
                "invoice_count": len(self.invoice_data)
            },
            "invoice_details": invoice_details,
            "risk_warnings": self.risk_warnings
        }

        return result

    def check_match_rate(self) -> Dict[str, Any]:
        """
        检查报关单与发票匹配率

        匹配率 = 匹配的报关单金额 / 报关单总金额 × 100%
        要求：≥98%

        Returns:
            匹配率检查结果
        """
        total_customs_amount = sum(d.fob_usd for d in self.customs_data)
        matched_amount = 0.0
        unmatched = []

        for customs in self.customs_data:
            matched = False
            for invoice in self.invoice_data:
                if customs.goods_name == invoice.goods_name:
                    matched = True
                    matched_amount += customs.fob_usd
                    break
            if not matched:
                unmatched.append({
                    "declaration_id": customs.declaration_id,
                    "goods_name": customs.goods_name,
                    "fob_usd": customs.fob_usd
                })

        match_rate = (matched_amount / total_customs_amount * 100) if total_customs_amount > 0 else 0

        result = {
            "match_rate": match_rate,
            "total_customs_amount": total_customs_amount,
            "matched_amount": matched_amount,
            "unmatched_items": unmatched,
            "passed": match_rate >= 98
        }

        if match_rate < 98:
            self.risk_warnings.append({
                "type": "match_rate",
                "level": "high",
                "message": f"报关单与发票匹配率 {match_rate:.2f}% 低于98%要求",
                "details": unmatched
            })

        return result

    def check_exchange_cost(self) -> Dict[str, Any]:
        """
        检查换汇成本

        换汇成本 = 采购不含税进价 / 出口美元离岸价
        合理区间：5-8

        Returns:
            换汇成本检查结果
        """
        results = []

        for customs in self.customs_data:
            # 查找对应的采购发票
            matched_invoice = None
            for invoice in self.invoice_data:
                if customs.goods_name == invoice.goods_name:
                    matched_invoice = invoice
                    break

            if matched_invoice:
                purchase_ex_tax = matched_invoice.amount_with_tax / (1 + matched_invoice.tax_rate / 100)
                purchase_cny = purchase_ex_tax
                export_usd = customs.fob_usd

                if export_usd > 0:
                    exchange_cost = purchase_cny / export_usd
                    is_normal = 5 <= exchange_cost <= 8

                    result = {
                        "declaration_id": customs.declaration_id,
                        "goods_name": customs.goods_name,
                        "purchase_cny": purchase_cny,
                        "export_usd": export_usd,
                        "exchange_cost": exchange_cost,
                        "is_normal": is_normal
                    }

                    if not is_normal:
                        level = "high" if exchange_cost < 5 else "medium"
                        self.risk_warnings.append({
                            "type": "exchange_cost",
                            "level": level,
                            "message": f"换汇成本 {exchange_cost:.2f} 超出合理区间(5-8)",
                            "details": result
                        })

                    results.append(result)

        return {
            "exchange_cost_details": results,
            "has_anomaly": any(not r["is_normal"] for r in results)
        }

    def run_full_calculation(self) -> Dict[str, Any]:
        """
        执行完整计算（包含风险检查）

        Returns:
            完整的计算结果报告
        """
        # 执行匹配率检查
        match_check = self.check_match_rate()

        # 执行换汇成本检查
        exchange_cost_check = self.check_exchange_cost()

        # 根据企业类型执行退税计算
        if self.enterprise_type == "production":
            rebate_calc = self.calculate_production_type()
        else:
            rebate_calc = self.calculate_trading_type()

        # 组装完整报告
        report = {
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "enterprise_type": rebate_calc["enterprise_type"],
            "customs_declarations": [asdict(d) for d in self.customs_data],
            "purchase_invoices": [asdict(inv) for inv in self.invoice_data],
            "match_check": match_check,
            "exchange_cost_check": exchange_cost_check,
            "rebate_calculation": rebate_calc,
            "total_rebate_amount": rebate_calc["calculation_details"].get("total_rebate_amount", 0),
            "risk_warnings": self.risk_warnings
        }

        return report


def load_data_from_json(json_file: str) -> Dict[str, Any]:
    """从JSON文件加载数据"""
    with open(json_file, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_report_to_json(report: Dict[str, Any], output_file: str):
    """保存报告到JSON文件"""
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)


def main():
    parser = argparse.ArgumentParser(description="出口退税计算脚本")
    parser.add_argument("--type", "-t", choices=["production", "trading"],
                        default="production", help="企业类型：production(生产型) 或 trading(外贸型)")
    parser.add_argument("--exchange-rate", "-e", type=float, required=True,
                        help="汇率（人民币/美元）")
    parser.add_argument("--data-file", "-d", help="输入数据JSON文件路径")
    parser.add_argument("--output", "-o", help="输出报告JSON文件路径")

    args = parser.parse_args()

    # 创建计算器
    calculator = ExportRebateCalculator(enterprise_type=args.type)
    calculator.set_exchange_rate(args.exchange_rate)

    # 如果提供了数据文件，加载数据
    if args.data_file:
        data = load_data_from_json(args.data_file)

        # 加载报关单数据
        for item in data.get("customs_declarations", []):
            calculator.add_customs_declaration(CustomsDeclaration(**item))

        # 加载采购发票数据
        for item in data.get("purchase_invoices", []):
            calculator.add_purchase_invoice(PurchaseInvoice(**item))

    # 执行计算
    report = calculator.run_full_calculation()

    # 输出报告
    print("\n" + "="*60)
    print("出口退税计算报告")
    print("="*60)
    print(f"\n企业类型：{report['enterprise_type']}")
    print(f"报关单数量：{len(report['customs_declarations'])}")
    print(f"采购发票数量：{len(report['purchase_invoices'])}")
    print(f"\n【匹配率检查】")
    print(f"  匹配率：{report['match_check']['match_rate']:.2f}%")
    print(f"  是否通过：{'是' if report['match_check']['passed'] else '否'}")

    print(f"\n【换汇成本检查】")
    for detail in report['exchange_cost_check'].get('exchange_cost_details', []):
        status = "正常" if detail['is_normal'] else "异常"
        print(f"  {detail['goods_name']}: {detail['exchange_cost']:.2f} ({status})")

    print(f"\n【退税计算结果】")
    rebate_details = report['rebate_calculation']['calculation_details']
    print(f"  出口离岸价(FOB)：{rebate_details.get('total_fob_usd', 0):,.2f} 美元")
    print(f"  折合人民币：{rebate_details.get('total_fob_cny', 0):,.2f} 元")
    print(f"  应退税额：{report['total_rebate_amount']:,.2f} 元")

    if report['risk_warnings']:
        print(f"\n【风险预警】({len(report['risk_warnings'])}项)")
        for warning in report['risk_warnings']:
            print(f"  [{warning['level'].upper()}] {warning['message']}")

    # 保存报告
    if args.output:
        save_report_to_json(report, args.output)
        print(f"\n报告已保存至：{args.output}")

    print("\n" + "="*60)


if __name__ == "__main__":
    main()
