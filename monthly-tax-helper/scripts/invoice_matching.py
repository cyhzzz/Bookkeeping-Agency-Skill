"""
进项发票与销项发票比对脚本
用于验证进项销项匹配度、税率一致性、农产品收购发票抵扣等

输入：进项发票列表、销项发票列表
输出：比对结果（正常/异常/预警）

比对规则：
- 进项销项金额匹配度应在95%-105%区间
- 进项税率应与销项对应商品税率一致
- 农产品收购发票：计算抵扣金额比对
"""

import argparse
import json
import sys
from datetime import datetime
from typing import Dict, List, Optional, Any, Tuple


# 税率标准映射表（商品类别 -> 标准税率）
TAX_RATE_STANDARDS = {
    "图书报纸": 9.0,
    "粮食农产品": 9.0,
    "自来水": 9.0,
    "暖气": 9.0,
    "石油液化气": 9.0,
    "天然气": 9.0,
    "沼 气": 9.0,
    "居民用煤炭制品": 9.0,
    "饲料": 9.0,
    "化肥": 9.0,
    "农药": 9.0,
    "农膜": 9.0,
    "农业产品": 9.0,
    "交通运输": 9.0,
    "基础电信": 9.0,
    "建筑": 9.0,
    "不动产租赁": 9.0,
    "销售不动产": 9.0,
    "邮政服务": 9.0,
    "电信服务-基础电信": 9.0,
    "电信服务-增值电信": 6.0,
    "研发技术": 6.0,
    "信息技术": 6.0,
    "文化创意": 6.0,
    "物流辅助": 6.0,
    "签证咨询": 6.0,
    "广播影视": 6.0,
    "保险服务": 6.0,
    "金融服务": 6.0,
    "生活服务": 6.0,
    "一般货物": 13.0,
    "劳务": 13.0,
    "有形动产租赁": 13.0
}


# 农产品收购发票抵扣率
AGRICULTURAL_PRODUCT_DEDUCTION_RATES = {
    "自开票一般纳税人": 9.0,  # 购进农产品按9%计算抵扣
    "小规模纳税人": 3.0,      # 购进农产品按3%计算抵扣（自开或代开）
    "加计扣除": 1.0          # 深加工13%税率货物，可再加计1%扣除
}


class Invoice:
    """发票数据类"""
    def __init__(self, invoice_data: Dict[str, Any]):
        self.id = invoice_data.get("id", "")
        self.type = invoice_data.get("type", "")
        self.code = invoice_data.get("code", "")
        self.number = invoice_data.get("number", "")
        self.date = invoice_data.get("date", "")
        self.buyer = invoice_data.get("buyer", {})
        self.seller = invoice_data.get("seller", {})
        self.items = invoice_data.get("items", [])
        self.amount = invoice_data.get("amount", 0.0)  # 不含税金额
        self.tax_rate = invoice_data.get("tax_rate", 0.0)
        self.tax_amount = invoice_data.get("tax_amount", 0.0)
        self.total_amount = invoice_data.get("total_amount", 0.0)  # 价税合计
        self.category = invoice_data.get("category", "")  # A/B/C/D/E类

    @property
    def is_input(self) -> bool:
        """是否为进项发票（采购发票）"""
        return self.category in ["B", "C"] and self.type in ["增值税专用发票", "增值税电子专用发票"]

    @property
    def is_output(self) -> bool:
        """是否为销项发票（销售发票）"""
        return self.category == "A" and self.type in ["增值税专用发票", "增值税普通发票", "增值税电子普通发票"]

    @property
    def is_agricultural(self) -> bool:
        """是否为农产品收购发票"""
        return "农产品" in str(self.items) or "农业产品" in str(self.items)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "type": self.type,
            "code": self.code,
            "number": self.number,
            "date": self.date,
            "buyer": self.buyer,
            "seller": self.seller,
            "amount": self.amount,
            "tax_rate": self.tax_rate,
            "tax_amount": self.tax_amount,
            "total_amount": self.total_amount,
            "category": self.category
        }


def validate_amount_match(input_invoices: List[Invoice], output_invoices: List[Invoice],
                         tolerance: float = 0.05) -> Dict[str, Any]:
    """
    验证进项销项金额匹配度

    Args:
        input_invoices: 进项发票列表
        output_invoices: 销项发票列表
        tolerance: 容许偏差（默认5%）

    Returns:
        匹配验证结果
    """
    total_input_amount = sum(inv.total_amount for inv in input_invoices)
    total_output_amount = sum(inv.total_amount for inv in output_invoices)

    if total_output_amount == 0:
        return {
            "check_type": "金额匹配度",
            "status": "warning",
            "message": "销项金额为0，无法进行匹配度计算",
            "details": {
                "total_input": total_input_amount,
                "total_output": total_output_amount
            }
        }

    match_ratio = total_input_amount / total_output_amount if total_output_amount != 0 else 0
    match_percentage = match_ratio * 100

    lower_bound = (1 - tolerance) * 100  # 95%
    upper_bound = (1 + tolerance) * 100  # 105%

    if lower_bound <= match_percentage <= upper_bound:
        status = "normal"
        message = f"进项销项匹配度正常，当前比例{match_percentage:.1f}%，在{lower_bound:.0f}%-{upper_bound:.0f}%区间内"
    elif match_percentage < lower_bound:
        status = "abnormal"
        message = f"进项金额偏低，当前比例{match_percentage:.1f}%，低于{lower_bound:.0f}%，可能存在进项抵扣不足"
    else:
        status = "warning"
        message = f"进项金额偏高，当前比例{match_percentage:.1f}%，高于{upper_bound:.0f}%，请确认是否有异常"

    return {
        "check_type": "金额匹配度",
        "status": status,
        "message": message,
        "details": {
            "total_input_amount": round(total_input_amount, 2),
            "total_output_amount": round(total_output_amount, 2),
            "match_ratio": round(match_percentage, 2),
            "tolerance": f"{tolerance*100:.0f}%",
            "lower_bound": lower_bound,
            "upper_bound": upper_bound
        }
    }


def validate_tax_rate_consistency(invoices: List[Invoice]) -> Dict[str, Any]:
    """
    验证发票税率与商品名称的一致性

    Args:
        invoices: 发票列表

    Returns:
        税率一致性验证结果
    """
    anomalies = []

    for inv in invoices:
        for item in inv.items:
            goods_name = item.get("name", "")
            actual_rate = inv.tax_rate * 100  # 转为百分比

            # 查找标准税率
            standard_rate = None
            for category, rate in TAX_RATE_STANDARDS.items():
                if category in goods_name:
                    standard_rate = rate
                    break

            if standard_rate is not None and abs(actual_rate - standard_rate) > 0.1:
                anomalies.append({
                    "invoice_id": inv.id,
                    "goods_name": goods_name,
                    "actual_rate": actual_rate,
                    "standard_rate": standard_rate,
                    "deviation": round(actual_rate - standard_rate, 2),
                    "severity": "high"
                })

    if not anomalies:
        return {
            "check_type": "税率一致性",
            "status": "normal",
            "message": "所有发票税率与商品名称匹配正确",
            "anomalies": []
        }

    return {
        "check_type": "税率一致性",
        "status": "abnormal",
        "message": f"发现{len(anomalies)}条税率异常记录",
        "anomalies": anomalies
    }


def validate_agricultural_deduction(input_invoices: List[Invoice],
                                   customer_type: str = "general") -> Dict[str, Any]:
    """
    验证农产品收购发票抵扣金额

    Args:
        input_invoices: 进项发票列表
        customer_type: 客户类型（general/small_scale）

    Returns:
        农产品抵扣验证结果
    """
    agricultural_invoices = [inv for inv in input_invoices if inv.is_agricultural]

    if not agricultural_invoices:
        return {
            "check_type": "农产品抵扣",
            "status": "normal",
            "message": "无农产品收购发票",
            "agricultural_invoices": []
        }

    deduction_rate = (AGRICULTURAL_PRODUCT_DEDUCTION_RATES.get(customer_type, 9.0) +
                      AGRICULTURAL_PRODUCT_DEDUCTION_RATES.get("加计扣除", 0))

    results = []
    total_deductible = 0
    total_actual_deductible = 0

    for inv in agricultural_invoices:
        # 计算理论抵扣金额（按不含税金额 * 抵扣率）
        theoretical_deduction = inv.amount * (deduction_rate / 100)
        # 实际抵扣金额（发票上的税额）
        actual_deductible = inv.tax_amount

        difference = theoretical_deduction - actual_deductible

        results.append({
            "invoice_id": inv.id,
            "invoice_date": inv.date,
            "goods_amount": inv.amount,
            "deduction_rate": deduction_rate,
            "theoretical_deduction": round(theoretical_deduction, 2),
            "actual_deductible": round(actual_deductible, 2),
            "difference": round(difference, 2),
            "status": "normal" if abs(difference) < 0.01 else "warning"
        })

        total_deductible += theoretical_deduction
        total_actual_deductible += actual_deductible

    all_normal = all(r["status"] == "normal" for r in results)

    return {
        "check_type": "农产品抵扣",
        "status": "normal" if all_normal else "warning",
        "message": f"共{len(agricultural_invoices)}张农产品收购发票，理论抵扣{total_deductible:.2f}元，实际抵扣{total_actual_deductible:.2f}元",
        "deduction_rate_used": deduction_rate,
        "customer_type": customer_type,
        "invoice_details": results,
        "total_theoretical_deduction": round(total_deductible, 2),
        "total_actual_deductible": round(total_actual_deductible, 2)
    }


def validate_duplicate_invoices(invoices: List[Invoice]) -> Dict[str, Any]:
    """
    检查重复发票

    Args:
        invoices: 发票列表

    Returns:
        重复发票检查结果
    """
    invoice_keys = {}
    duplicates = []

    for inv in invoices:
        key = f"{inv.code}-{inv.number}"
        if key in invoice_keys:
            duplicates.append({
                "invoice_id": inv.id,
                "code": inv.code,
                "number": inv.number,
                "date": inv.date,
                "amount": inv.total_amount,
                "first_seen_id": invoice_keys[key]
            })
        else:
            invoice_keys[key] = inv.id

    if not duplicates:
        return {
            "check_type": "重复发票",
            "status": "normal",
            "message": "未发现重复发票",
            "duplicates": []
        }

    return {
        "check_type": "重复发票",
        "status": "abnormal",
        "message": f"发现{len(duplicates)}组重复发票",
        "duplicates": duplicates
    }


def validate_input_output_match(input_file: str, output_file: str,
                                customer_type: str = "general") -> Dict[str, Any]:
    """
    进项销项发票比对主函数

    Args:
        input_file: 进项发票JSON文件路径
        output_file: 销项发票JSON文件路径
        customer_type: 客户类型

    Returns:
        完整的比对报告
    """
    # 读取发票数据
    with open(input_file, "r", encoding="utf-8") as f:
        input_data = json.load(f)

    with open(output_file, "r", encoding="utf-8") as f:
        output_data = json.load(f)

    # 转换为Invoice对象
    input_invoices = [Invoice(inv) for inv in input_data]
    output_invoices = [Invoice(inv) for inv in output_data]

    # 执行各项检查
    amount_check = validate_amount_match(input_invoices, output_invoices)
    input_rate_check = validate_tax_rate_consistency(input_invoices)
    output_rate_check = validate_tax_rate_consistency(output_invoices)
    agricultural_check = validate_agricultural_deduction(input_invoices, customer_type)
    input_duplicate_check = validate_duplicate_invoices(input_invoices)
    output_duplicate_check = validate_duplicate_invoices(output_invoices)

    # 汇总结果
    checks = [
        amount_check,
        input_rate_check,
        output_rate_check,
        agricultural_check,
        input_duplicate_check,
        output_duplicate_check
    ]

    abnormal_checks = [c for c in checks if c["status"] == "abnormal"]
    warning_checks = [c for c in checks if c["status"] == "warning"]

    if abnormal_checks:
        overall_status = "abnormal"
        overall_message = f"发现{abnormal_checks.__len__()}项异常，需要人工处理"
    elif warning_checks:
        overall_status = "warning"
        overall_message = f"发现{warning_checks.__len__()}项预警，请关注"
    else:
        overall_status = "normal"
        overall_message = "所有比对项目正常"

    report = {
        "report_info": {
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "customer_type": customer_type,
            "input_invoice_count": len(input_invoices),
            "output_invoice_count": len(output_invoices)
        },
        "overall_status": overall_status,
        "overall_message": overall_message,
        "checks": checks,
        "summary": {
            "total_checks": len(checks),
            "normal_count": len([c for c in checks if c["status"] == "normal"]),
            "warning_count": len(warning_checks),
            "abnormal_count": len(abnormal_checks)
        },
        "recommendations": []
    }

    # 添加建议
    if amount_check["status"] != "normal":
        report["recommendations"].append(amount_check["message"])

    if input_rate_check["status"] != "normal":
        report["recommendations"].append("存在税率异常的进项发票，请核实商品分类是否正确")

    if agricultural_check["status"] != "normal":
        report["recommendations"].append("农产品收购发票抵扣金额存在差异，请核实抵扣率是否正确")

    if input_duplicate_check["status"] != "normal" or output_duplicate_check["status"] != "normal":
        report["recommendations"].append("存在重复发票，请核实是否重复入账")

    return report


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description="进项销项发票比对脚本")
    parser.add_argument("--input", required=True, help="进项发票JSON文件路径")
    parser.add_argument("--output", required=True, help="销项发票JSON文件路径")
    parser.add_argument("--customer-type", default="general",
                       choices=["general", "small_scale"],
                       help="客户类型")
    parser.add_argument("--report-output", help="比对报告输出路径（JSON格式）")

    args = parser.parse_args()

    # 执行比对
    result = validate_input_output_match(args.input, args.output, args.customer_type)

    # 输出结果
    if args.report_output:
        with open(args.report_output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"比对报告已保存至: {args.report_output}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
