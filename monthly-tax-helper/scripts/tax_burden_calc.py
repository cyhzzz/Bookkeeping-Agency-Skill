"""
税负率计算脚本
用于计算增值税税负率、企业所得税税负率和综合税负率

输入：客户ID、申报月份
输出：各税种税负率、行业均值对比、偏离度分析

计算公式：
- 增值税税负率 = 当期应纳税额 / 当期应税销售额 × 100%
- 企业所得税税负率 = 当期应纳税所得额 / 收入总额 × 100%
- 综合税负率 = 所有税种税额合计 / 营业收入 × 100%
"""

import argparse
import json
import sys
from datetime import datetime
from typing import Dict, List, Optional, Any


# 行业税负率均值参考数据（按行业分类）
INDUSTRY_BENCHMARKS = {
    "general": {  # 一般纳税人
        "增值税": 3.5,
        "企业所得税": 2.0,
        "综合税负率": 5.5
    },
    "small_scale": {  # 小规模纳税人
        "增值税": 1.0,
        "企业所得税": 2.0,
        "综合税负率": 3.0
    },
    "service": {  # 服务业
        "增值税": 3.0,
        "企业所得税": 2.5,
        "综合税负率": 5.5
    },
    "trade": {  # 商贸
        "增值税": 2.5,
        "企业所得税": 1.8,
        "综合税负率": 4.3
    },
    "manufacturing": {  # 制造业
        "增值税": 3.2,
        "企业所得税": 2.0,
        "综合税负率": 5.2
    },
    "construction": {  # 建筑业
        "增值税": 3.8,
        "企业所得税": 2.2,
        "综合税负率": 6.0
    },
    "technology": {  # 科技型企业
        "增值税": 2.8,
        "企业所得税": 1.5,
        "综合税负率": 4.3
    }
}


def calculate_vat_burden_rate(taxable_sales: float, vat_payable: float) -> Dict[str, Any]:
    """
    计算增值税税负率

    Args:
        taxable_sales: 当期应税销售额（不含税）
        vat_payable: 当期应纳税额（增值税）

    Returns:
        包含税负率和分析结果的字典
    """
    if taxable_sales == 0:
        return {
            "tax_type": "增值税",
            "tax_amount": vat_payable,
            "taxable_sales": taxable_sales,
            "tax_rate": None,
            "error": "销售额为0，无法计算税负率"
        }

    tax_rate = (vat_payable / taxable_sales) * 100

    return {
        "tax_type": "增值税",
        "tax_amount": round(vat_payable, 2),
        "taxable_sales": round(taxable_sales, 2),
        "tax_rate": round(tax_rate, 2),
        "unit": "%"
    }


def calculate_cit_burden_rate(total_income: float, taxable_income: float, cit_payable: float) -> Dict[str, Any]:
    """
    计算企业所得税税负率

    Args:
        total_income: 收入总额
        taxable_income: 应纳税所得额
        cit_payable: 企业所得税应纳税额

    Returns:
        包含税负率和分析结果的字典
    """
    if total_income == 0:
        return {
            "tax_type": "企业所得税",
            "tax_amount": cit_payable,
            "taxable_income": taxable_income,
            "total_income": total_income,
            "tax_rate": None,
            "error": "收入为0，无法计算税负率"
        }

    tax_rate = (taxable_income / total_income) * 100

    return {
        "tax_type": "企业所得税",
        "tax_amount": round(cit_payable, 2),
        "taxable_income": round(taxable_income, 2),
        "total_income": round(total_income, 2),
        "tax_rate": round(tax_rate, 2),
        "unit": "%"
    }


def calculate_comprehensive_burden_rate(total_revenue: float, total_tax: float) -> Dict[str, Any]:
    """
    计算综合税负率

    Args:
        total_revenue: 营业收入总额
        total_tax: 所有税种税额合计

    Returns:
        包含综合税负率的字典
    """
    if total_revenue == 0:
        return {
            "tax_type": "综合税负率",
            "total_tax": total_tax,
            "total_revenue": total_revenue,
            "tax_rate": None,
            "error": "营业收入为0，无法计算税负率"
        }

    tax_rate = (total_tax / total_revenue) * 100

    return {
        "tax_type": "综合税负率",
        "total_tax": round(total_tax, 2),
        "total_revenue": round(total_revenue, 2),
        "tax_rate": round(tax_rate, 2),
        "unit": "%"
    }


def compare_with_industry(tax_rate: float, industry_rate: float) -> Dict[str, Any]:
    """
    与行业均值对比

    Args:
        tax_rate: 企业税负率
        industry_rate: 行业均值

    Returns:
        包含对比结果的字典
    """
    deviation = tax_rate - industry_rate
    deviation_rate = (deviation / industry_rate) * 100 if industry_rate != 0 else 0

    if deviation < -0.5:
        level = "过低"
        risk_level = "warning"
        risk_message = "税负率低于行业均值，可能存在税务风险"
    elif deviation > 1.0:
        level = "过高"
        risk_level = "info"
        risk_message = "税负率高于行业均值，可关注税务筹划空间"
    else:
        level = "正常"
        risk_level = "normal"
        risk_message = "税负率处于正常区间"

    return {
        "company_rate": tax_rate,
        "industry_rate": industry_rate,
        "deviation": round(deviation, 2),
        "deviation_rate": round(deviation_rate, 2),
        "level": level,
        "risk_level": risk_level,
        "risk_message": risk_message
    }


def generate_warnings(vat_rate: float, cit_rate: float, comprehensive_rate: float,
                      industry_type: str, customer_type: str) -> List[Dict[str, Any]]:
    """
    生成预警信息

    Args:
        vat_rate: 增值税税负率
        cit_rate: 企业所得税税负率
        comprehensive_rate: 综合税负率
        industry_type: 行业类型
        customer_type: 客户类型（一般纳税人/小规模纳税人）

    Returns:
        预警信息列表
    """
    warnings = []
    benchmarks = INDUSTRY_BENCHMARKS.get(industry_type, INDUSTRY_BENCHMARKS["general"])

    # 增值税预警
    if vat_rate is not None:
        vat_deviation = vat_rate - benchmarks["增值税"]
        if vat_deviation < -1.0:
            warnings.append({
                "code": "W_VAT_001",
                "tax_type": "增值税",
                "level": "high",
                "message": f"增值税税负率{vat_rate}%低于行业均值{benchmarks['增值税']}%，偏低超过1个百分点",
                "suggestion": "请确认是否存在留抵税额、进货退回或税收优惠政策享受情况"
            })
        elif vat_deviation > 2.0:
            warnings.append({
                "code": "W_VAT_002",
                "tax_type": "增值税",
                "level": "medium",
                "message": f"增值税税负率{vat_rate}%高于行业均值{benchmarks['增值税']}%，偏高超过2个百分点",
                "suggestion": "可关注进项税额抵扣是否充分，是否有合理的税务筹划空间"
            })

    # 企业所得税预警
    if cit_rate is not None:
        cit_deviation = cit_rate - benchmarks["企业所得税"]
        if cit_deviation < -0.5:
            warnings.append({
                "code": "W_CIT_001",
                "tax_type": "企业所得税",
                "level": "high",
                "message": f"企业所得税税负率{cit_rate}%低于行业均值{benchmarks['企业所得税']}%，偏低超过0.5个百分点",
                "suggestion": "请确认成本费用列支是否合规，是否存在调增风险"
            })
        elif cit_deviation > 1.5:
            warnings.append({
                "code": "W_CIT_002",
                "tax_type": "企业所得税",
                "level": "medium",
                "message": f"企业所得税税负率{cit_rate}%高于行业均值{benchmarks['企业所得税']}%，偏高超过1.5个百分点",
                "suggestion": "可关注成本费用是否充分列支，是否有合理的税前扣除项目"
            })

    # 综合税负率预警
    if comprehensive_rate is not None:
        comp_deviation = comprehensive_rate - benchmarks["综合税负率"]
        if comp_deviation < -2.0:
            warnings.append({
                "code": "W_COMP_001",
                "tax_type": "综合税负率",
                "level": "high",
                "message": f"综合税负率{comprehensive_rate}%低于行业均值{benchmarks['综合税负率']}%，偏离较大",
                "suggestion": "请进行全面税务风险排查"
            })

    return warnings


def calculate_tax_burden(customer_id: str, filing_month: str,
                         taxable_sales: float, vat_payable: float,
                         total_income: float, taxable_income: float, cit_payable: float,
                         total_revenue: float, total_tax: float,
                         industry_type: str = "general",
                         customer_type: str = "general") -> Dict[str, Any]:
    """
    计算企业税负率（主函数）

    Args:
        customer_id: 客户ID
        filing_month: 申报月份（YYYY-MM格式）
        taxable_sales: 当期应税销售额（增值税计税依据）
        vat_payable: 当期应纳增值税税额
        total_income: 收入总额（企业所得税）
        taxable_income: 应纳税所得额（企业所得税）
        cit_payable: 当期应纳企业所得税税额
        total_revenue: 营业收入总额
        total_tax: 所有税种税额合计
        industry_type: 行业类型
        customer_type: 客户类型（general/small_scale）

    Returns:
        完整的税负率计算报告
    """
    # 计算各税种税负率
    vat_result = calculate_vat_burden_rate(taxable_sales, vat_payable)
    cit_result = calculate_cit_burden_rate(total_income, taxable_income, cit_payable)
    comprehensive_result = calculate_comprehensive_burden_rate(total_revenue, total_tax)

    # 获取行业均值
    industry_key = customer_type if customer_type in ["general", "small_scale"] else "general"
    benchmarks = INDUSTRY_BENCHMARKS.get(industry_key, INDUSTRY_BENCHMARKS["general"])

    # 行业对比
    vat_comparison = compare_with_industry(vat_result.get("tax_rate"), benchmarks["增值税"]) if vat_result.get("tax_rate") is not None else None
    cit_comparison = compare_with_industry(cit_result.get("tax_rate"), benchmarks["企业所得税"]) if cit_result.get("tax_rate") is not None else None
    comprehensive_comparison = compare_with_industry(comprehensive_result.get("tax_rate"), benchmarks["综合税负率"]) if comprehensive_result.get("tax_rate") is not None else None

    # 生成预警
    warnings = generate_warnings(
        vat_result.get("tax_rate"),
        cit_result.get("tax_rate"),
        comprehensive_result.get("tax_rate"),
        industry_type,
        customer_type
    )

    # 组装结果
    report = {
        "report_info": {
            "customer_id": customer_id,
            "filing_month": filing_month,
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "industry_type": industry_type,
            "customer_type": customer_type
        },
        "tax_burden_details": {
            "增值税": vat_result,
            "企业所得税": cit_result,
            "综合税负率": comprehensive_result
        },
        "industry_comparison": {
            "增值税": vat_comparison,
            "企业所得税": cit_comparison,
            "综合税负率": comprehensive_comparison
        },
        "warnings": warnings,
        "summary": {
            "overall_assessment": "正常" if len([w for w in warnings if w["level"] == "high"]) == 0 else "需关注",
            "key_concerns": [w["message"] for w in warnings if w["level"] == "high"],
            "suggestions": list(set([w["suggestion"] for w in warnings]))
        }
    }

    return report


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description="税负率计算脚本")
    parser.add_argument("--customer-id", required=True, help="客户ID")
    parser.add_argument("--filing-month", required=True, help="申报月份（YYYY-MM格式）")
    parser.add_argument("--taxable-sales", type=float, required=True, help="应税销售额（不含税）")
    parser.add_argument("--vat-payable", type=float, required=True, help="应纳增值税税额")
    parser.add_argument("--total-income", type=float, required=True, help="收入总额")
    parser.add_argument("--taxable-income", type=float, required=True, help="应纳税所得额")
    parser.add_argument("--cit-payable", type=float, required=True, help="应纳企业所得税税额")
    parser.add_argument("--total-revenue", type=float, required=True, help="营业收入总额")
    parser.add_argument("--total-tax", type=float, required=True, help="所有税种税额合计")
    parser.add_argument("--industry-type", default="general",
                       choices=["general", "small_scale", "service", "trade", "manufacturing", "construction", "technology"],
                       help="行业类型")
    parser.add_argument("--customer-type", default="general",
                       choices=["general", "small_scale"],
                       help="客户类型")
    parser.add_argument("--output", help="输出文件路径（JSON格式）")

    args = parser.parse_args()

    # 计算税负率
    result = calculate_tax_burden(
        customer_id=args.customer_id,
        filing_month=args.filing_month,
        taxable_sales=args.taxable_sales,
        vat_payable=args.vat_payable,
        total_income=args.total_income,
        taxable_income=args.taxable_income,
        cit_payable=args.cit_payable,
        total_revenue=args.total_revenue,
        total_tax=args.total_tax,
        industry_type=args.industry_type,
        customer_type=args.customer_type
    )

    # 输出结果
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"报告已保存至: {args.output}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
