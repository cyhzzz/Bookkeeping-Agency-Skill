"""
研发费用加计扣除计算脚本

2023年起适用比例：100%

输入：
- 研发项目列表
- 各项目费用明细（人员人工/直接投入/折旧/无形资产/其他）

输出：
- 研发费用辅助账汇总表
- 加计扣除金额计算表
- 优惠备案建议

费用归集规则：
- 人员人工：研发人员工资社保（需劳动合同+社保记录）
- 直接投入：材料、水电、测试费等（需发票）
- 折旧费用：研发设备折旧（需固定资产卡片）
- 无形资产摊销：专利、软件摊销
- 其他费用：不超过可扣除总额的10%
"""

import json
import argparse
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class RDProject:
    """研发项目信息"""
    project_id: str  # 项目编号
    project_name: str  # 项目名称
    start_date: str  # 开始日期
    end_date: str  # 结束日期
    status: str  # 研发中/已完成


@dataclass
class RDExpense:
    """研发费用明细"""
    project_id: str  # 项目编号
    expense_category: str  # 费用类别
    amount: float  # 金额
    description: str  # 说明
    voucher_id: str  # 凭证号
    expense_date: str  # 日期


class RDExpenseCalculator:
    """研发费用加计扣除计算器"""

    # 费用类别
    CATEGORIES = [
        "人员人工",
        "直接投入",
        "折旧费用",
        "无形资产摊销",
        "其他费用"
    ]

    # 2023年最新政策
    SUPER_DEDUCTION_RATE = 100  # 加计扣除比例
    CIT_RATE = 25  # 企业所得税税率（法定）
    HIGH_TECH_CIT_RATE = 15  # 高新技术企业优惠税率

    # 其他费用上限比例
    OTHER_EXPENSE_LIMIT = 10

    # 负面清单行业
    NEGATIVE_LIST = [
        "烟草制造业",
        "住宿和餐饮业",
        "批发和零售业",
        "房地产业",
        "租赁和商务服务业",
        "娱乐业"
    ]

    def __init__(self):
        self.projects: List[RDProject] = []
        self.expenses: List[RDExpense] = []
        self.industry: str = ""

    def add_project(self, project: RDProject):
        """添加研发项目"""
        self.projects.append(project)

    def add_expense(self, expense: RDExpense):
        """添加研发费用"""
        self.expenses.append(expense)

    def set_industry(self, industry: str):
        """设置所属行业"""
        self.industry = industry

    def is_in_negative_list(self) -> bool:
        """检查是否在负面清单"""
        return self.industry in self.NEGATIVE_LIST

    def calculate_by_project(self) -> Dict[str, Any]:
        """
        按项目汇总研发费用

        Returns:
            各项目费用汇总
        """
        project_expenses = {}

        for project in self.projects:
            project_expenses[project.project_id] = {
                "project_id": project.project_id,
                "project_name": project.project_name,
                "status": project.status,
                "expenses": {cat: 0.0 for cat in self.CATEGORIES},
                "total": 0.0
            }

        for expense in self.expenses:
            if expense.project_id in project_expenses:
                cat = expense.expense_category
                if cat in project_expenses[expense.project_id]["expenses"]:
                    project_expenses[expense.project_id]["expenses"][cat] += expense.amount
                    project_expenses[expense.project_id]["total"] += expense.amount

        return project_expenses

    def calculate_summary(self) -> Dict[str, Any]:
        """
        计算研发费用汇总及加计扣除

        Returns:
            汇总计算结果
        """
        project_expenses = self.calculate_by_project()

        # 计算各类别合计
        category_totals = {cat: 0.0 for cat in self.CATEGORIES}
        grand_total = 0.0

        for pid, data in project_expenses.items():
            for cat, amount in data["expenses"].items():
                category_totals[cat] += amount
            grand_total += data["total"]

        # 检查其他费用是否超限
        other_ratio = (category_totals["其他费用"] / grand_total * 100) if grand_total > 0 else 0
        other_expense_valid = other_ratio <= self.OTHER_EXPENSE_LIMIT

        # 计算可加计扣除金额
        if self.is_in_negative_list():
            super_deduction_amount = 0.0
            is_eligible = False
            reason = f"所属行业'{self.industry}'在负面清单内，不适用加计扣除政策"
        else:
            super_deduction_amount = grand_total * (self.SUPER_DEDUCTION_RATE / 100)
            is_eligible = True
            reason = "符合加计扣除条件"

        # 计算节税金额
        tax_saving_standard = super_deduction_amount * (self.CIT_RATE / 100)
        tax_saving_high_tech = super_deduction_amount * (self.HIGH_TECH_CIT_RATE / 100)

        result = {
            "category_totals": category_totals,
            "grand_total": grand_total,
            "other_expense_ratio": other_ratio,
            "other_expense_valid": other_expense_valid,
            "super_deduction_rate": self.SUPER_DEDUCTION_RATE,
            "super_deduction_amount": super_deduction_amount,
            "is_eligible": is_eligible,
            "reason": reason,
            "tax_saving_standard": tax_saving_standard,  # 法定税率下节税
            "tax_saving_high_tech": tax_saving_high_tech,  # 高新企业税率下节税
            "projects": project_expenses
        }

        return result

    def generate_auxiliary_account(self) -> List[Dict[str, Any]]:
        """
        生成研发费用辅助账

        Returns:
            辅助账明细列表
        """
        auxiliary = []

        for expense in self.expenses:
            project_name = ""
            for proj in self.projects:
                if proj.project_id == expense.project_id:
                    project_name = proj.project_name
                    break

            auxiliary.append({
                "project_id": expense.project_id,
                "project_name": project_name,
                "expense_category": expense.expense_category,
                "amount": expense.amount,
                "description": expense.description,
                "voucher_id": expense.voucher_id,
                "expense_date": expense.expense_date
            })

        # 按项目和日期排序
        auxiliary.sort(key=lambda x: (x["project_id"], x["expense_date"]))

        return auxiliary

    def generate_filing_suggestions(self, summary: Dict[str, Any]) -> Dict[str, Any]:
        """
        生成优惠备案建议

        Args:
            summary: 费用汇总结果

        Returns:
            备案建议
        """
        suggestions = {
            "is_eligible": summary["is_eligible"],
            "super_deduction_amount": summary["super_deduction_amount"],
            "filing_deadlines": {
                "pre_payment": "每季度结束后15日内",
                "annual_settlement": "次年5月31日前"
            },
            "required_documents": [
                "研发项目立项文件",
                "研发人员名单及工资表",
                "研发费用辅助账",
                "研发设备折旧计算表",
                "无形资产摊销计算表",
                "委托研发合同（如有）",
                "会计凭证及原始凭证"
            ],
            "applicable_policy": "《财政部 税务总局关于进一步完善研发费用税前加计扣除政策的公告》（2023年第7号）",
            "notes": []
        }

        # 添加注意事项
        if not summary["other_expense_valid"]:
            suggestions["notes"].append(
                f"其他费用占比 {summary['other_expense_ratio']:.1f}% 超过10%上限，需调整"
            )

        if not summary["is_eligible"]:
            suggestions["notes"].append(
                f"企业所属行业 '{self.industry}' 在负面清单内，不适用加计扣除政策"
            )

        return suggestions

    def run_full_calculation(self, year: int = 2024) -> Dict[str, Any]:
        """
        执行完整计算

        Args:
            year: 所属年度

        Returns:
            完整报告
        """
        summary = self.calculate_summary()
        auxiliary = self.generate_auxiliary_account()
        filing = self.generate_filing_suggestions(summary)

        # 按项目汇总
        project_summary = []
        for pid, data in summary["projects"].items():
            project_summary.append({
                "project_id": pid,
                "project_name": data["project_name"],
                "status": data["status"],
                "total_expense": data["total"]
            })

        report = {
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "report_year": year,
            "industry": self.industry,
            "is_in_negative_list": self.is_in_negative_list(),
            "summary": {
                "category_totals": summary["category_totals"],
                "grand_total": summary["grand_total"],
                "other_expense_ratio": summary["other_expense_ratio"],
                "other_expense_valid": summary["other_expense_valid"]
            },
            "super_deduction": {
                "rate": summary["super_deduction_rate"],
                "amount": summary["super_deduction_amount"],
                "is_eligible": summary["is_eligible"],
                "reason": summary["reason"]
            },
            "tax_saving": {
                "standard_rate": summary["tax_saving_standard"],
                "high_tech_rate": summary["tax_saving_high_tech"]
            },
            "projects": project_summary,
            "auxiliary_account": auxiliary,
            "filing_suggestions": filing
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
    parser = argparse.ArgumentParser(description="研发费用加计扣除计算脚本")
    parser.add_argument("--year", "-y", type=int, default=2024, help="所属年度")
    parser.add_argument("--industry", "-i", default="制造业", help="所属行业")
    parser.add_argument("--data-file", "-d", help="输入数据JSON文件路径")
    parser.add_argument("--output", "-o", help="输出报告JSON文件路径")

    args = parser.parse_args()

    # 创建计算器
    calculator = RDExpenseCalculator()
    calculator.set_industry(args.industry)

    # 如果提供了数据文件，加载数据
    if args.data_file:
        data = load_data_from_json(args.data_file)

        # 加载项目数据
        for item in data.get("projects", []):
            calculator.add_project(RDProject(**item))

        # 加载费用数据
        for item in data.get("expenses", []):
            calculator.add_expense(RDExpense(**item))

    # 执行计算
    report = calculator.run_full_calculation(args.year)

    # 输出报告
    print("\n" + "="*60)
    print(f"研发费用加计扣除计算报告（{args.year}年度）")
    print("="*60)
    print(f"\n所属行业：{report['industry']}")
    print(f"是否在负面清单：{'是' if report['is_in_negative_list'] else '否'}")

    print(f"\n【研发费用汇总】")
    for cat, amount in report['summary']['category_totals'].items():
        print(f"  {cat}：{amount:,.2f} 元")
    print(f"  合计：{report['summary']['grand_total']:,.2f} 元")

    print(f"\n【其他费用占比检查】")
    print(f"  占比：{report['summary']['other_expense_ratio']:.2f}%")
    print(f"  是否合规：{'是' if report['summary']['other_expense_valid'] else '否'}")

    print(f"\n【加计扣除计算】")
    deduction = report['super_deduction']
    print(f"  加计扣除比例：{deduction['rate']}%")
    print(f"  可加计扣除金额：{deduction['amount']:,.2f} 元")
    print(f"  是否符合条件：{'是' if deduction['is_eligible'] else '否'}")
    if not deduction['is_eligible']:
        print(f"  原因：{deduction['reason']}")

    print(f"\n【节税测算】")
    tax_saving = report['tax_saving']
    print(f"  法定税率(25%)下节约税款：{tax_saving['standard_rate']:,.2f} 元")
    print(f"  高新企业税率(15%)下节约税款：{tax_saving['high_tech_rate']:,.2f} 元")

    print(f"\n【研发项目】")
    for proj in report['projects']:
        print(f"  {proj['project_name']}（{proj['project_id']}）：{proj['total_expense']:,.2f} 元")

    print(f"\n【备案建议】")
    filing = report['filing_suggestions']
    print(f"  季度预缴申报：{filing['filing_deadlines']['pre_payment']}")
    print(f"  年度汇算清缴：{filing['filing_deadlines']['annual_settlement']}")
    print(f"  适用政策：{filing['applicable_policy']}")

    # 保存报告
    if args.output:
        save_report_to_json(report, args.output)
        print(f"\n报告已保存至：{args.output}")

    print("\n" + "="*60)


if __name__ == "__main__":
    main()
