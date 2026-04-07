"""
跨区域经营企业所得税分摊计算

支持总分公司、跨省市分摊

输入：
- 总机构基本信息
- 各分公司信息（收入、工资、资产）
- 各省市职工工资比例

输出：
- 总分机构所得税分摊表
- 各地应纳税额
- 地方优惠政策的适用性判断
"""

import json
import argparse
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class Headquarters:
    """总机构信息"""
    name: str  # 总机构名称
    location: str  # 所在地
    tax_type: str  # 纳税人类型（一般纳税人/小规模纳税人）
    total_income: float  # 总收入
    total_payroll: float  # 职工薪酬总额
    total_assets: float  # 资产总额
    applicable_rate: float  # 适用税率
    is_high_tech: bool  # 是否高新技术企业


@dataclass
class Branch:
    """分支机构信息"""
    branch_id: str  # 分支机构编号
    name: str  # 分支机构名称
    location: str  # 所在地（省/市）
    income: float  # 营业收入
    payroll: float  # 职工薪酬
    assets: float  # 资产
    is_local_assessment: bool  # 是否就地预缴
    applicable_policies: List[str]  # 适用的当地优惠政策


class CrossRegionTaxAllocator:
    """跨区域所得税分摊计算器"""

    # 分摊权重（可调整）
    WEIGHT_INCOME = 0.5  # 收入权重
    WEIGHT_PAYROLL = 0.5  # 工资权重

    # 分摊比例（固定）
    HQ_SHARE = 50  # 总机构分摊比例
    BRANCH_SHARE = 50  # 分支机构分摊比例

    # 地区优惠政策汇总
    REGIONAL_POLICIES = {
        "西部地区": {
            "provinces": ["重庆", "四川", "贵州", "云南", "西藏", "陕西", "甘肃",
                         "青海", "宁夏", "新疆", "内蒙", "广西"],
            "policy": "西部大开发15%优惠税率",
            "condition": "符合产业目录"
        },
        "东北地区": {
            "provinces": ["辽宁", "吉林", "黑龙江"],
            "policy": "装备制造业等15%优惠",
            "condition": "特定产业"
        },
        "海南自贸港": {
            "provinces": ["海南"],
            "policy": "2025年前鼓励类15%；2035年实质性运营全部15%",
            "condition": "实质性运营"
        },
        "粤港澳大湾区": {
            "provinces": ["广州", "深圳", "珠海", "东莞", "惠州", "佛山", "中山", "江门", "肇庆"],
            "policy": "高新技术企业15%",
            "condition": "高新技术企业"
        },
        "长三角示范区": {
            "provinces": ["上海", "苏州", "嘉兴"],
            "policy": "高新技术企业15%",
            "condition": "高新技术企业"
        },
        "雄安新区": {
            "provinces": ["雄安"],
            "policy": "符合产业目录15%",
            "condition": "符合产业目录"
        }
    }

    def __init__(self):
        self.headquarters: Optional[Headquarters] = None
        self.branches: List[Branch] = []
        self.total_tax: float = 0.0

    def set_headquarters(self, hq: Headquarters):
        """设置总机构信息"""
        self.headquarters = hq

    def add_branch(self, branch: Branch):
        """添加分支机构"""
        self.branches.append(branch)

    def set_total_tax(self, tax: float):
        """设置总应纳所得税额"""
        self.total_tax = tax

    def identify_region(self, location: str) -> Optional[str]:
        """识别所属区域"""
        for region, info in self.REGIONAL_POLICIES.items():
            for prov in info["provinces"]:
                if prov in location:
                    return region
        return None

    def calculate_allocation_ratio(self) -> Dict[str, float]:
        """
        计算各分支机构分摊比例

        分摊比例 = (分支机构收入/总机构收入) × 权重1
                 + (分支机构工资/总机构工资) × 权重2

        Returns:
            各分支机构分摊比例
        """
        if not self.headquarters:
            raise ValueError("未设置总机构信息")

        total_income = self.headquarters.total_income
        total_payroll = self.headquarters.total_payroll

        ratios = {}

        for branch in self.branches:
            income_ratio = (branch.income / total_income * self.WEIGHT_INCOME) if total_income > 0 else 0
            payroll_ratio = (branch.payroll / total_payroll * self.WEIGHT_PAYROLL) if total_payroll > 0 else 0
            ratios[branch.branch_id] = income_ratio + payroll_ratio

        return ratios

    def calculate_allocation(self) -> Dict[str, Any]:
        """
        计算所得税分摊

        Returns:
            分摊计算结果
        """
        if not self.headquarters:
            raise ValueError("未设置总机构信息")

        # 计算各分支机构分摊比例
        ratios = self.calculate_allocation_ratio()

        # 计算税款分摊
        hq_tax = self.total_tax * (self.HQ_SHARE / 100)

        branch_allocations = []
        total_branch_tax = 0.0

        for branch in self.branches:
            ratio = ratios[branch.branch_id]
            branch_tax = self.total_tax * (self.BRANCH_SHARE / 100) * ratio

            # 计算当地应纳税额（考虑地方优惠）
            local_tax = branch_tax

            # 识别区域政策
            region = self.identify_region(branch.location)

            branch_allocations.append({
                "branch_id": branch.branch_id,
                "branch_name": branch.name,
                "location": branch.location,
                "region": region,
                "income": branch.income,
                "payroll": branch.payroll,
                "allocation_ratio": ratio * 100,
                "pre_tax_amount": branch_tax,
                "local_tax_amount": local_tax,
                "applicable_policies": branch.applicable_policies,
                "is_local_assessment": branch.is_local_assessment
            })

            total_branch_tax += branch_tax

        result = {
            "total_tax": self.total_tax,
            "hq_allocation": {
                "name": self.headquarters.name,
                "location": self.headquarters.location,
                "ratio": self.HQ_SHARE,
                "tax_amount": hq_tax,
                "is_high_tech": self.headquarters.is_high_tech,
                "applicable_rate": self.headquarters.applicable_rate
            },
            "branch_allocations": branch_allocations,
            "total_branch_tax": total_branch_tax,
            "verification": {
                "hq_branch_sum": hq_tax + total_branch_tax,
                "differs_from_total": abs(hq_tax + total_branch_tax - self.total_tax) > 0.01
            }
        }

        return result

    def analyze_regional_policies(self) -> Dict[str, Any]:
        """
        分析各地政策适用性

        Returns:
            政策分析结果
        """
        if not self.headquarters:
            raise ValueError("未设置总机构信息")

        policy_analysis = {
            "hq_location": self.headquarters.location,
            "hq_region": self.identify_region(self.headquarters.location),
            "branch_policies": []
        }

        # 总机构所在地政策
        if policy_analysis["hq_region"]:
            region_info = self.REGIONAL_POLICIES[policy_analysis["hq_region"]]
            policy_analysis["hq_policy"] = {
                "region": policy_analysis["hq_region"],
                "policy": region_info["policy"],
                "condition": region_info["condition"]
            }

        # 各分支机构政策
        for branch in self.branches:
            region = self.identify_region(branch.location)
            branch_policy = {
                "branch_name": branch.name,
                "location": branch.location,
                "region": region
            }

            if region:
                region_info = self.REGIONAL_POLICIES[region]
                branch_policy["applicable"] = {
                    "policy": region_info["policy"],
                    "condition": region_info["condition"],
                    "meets_condition": len(branch.applicable_policies) > 0 or region == "海南自贸港"
                }
            else:
                branch_policy["applicable"] = {
                    "policy": "无特殊优惠政策",
                    "condition": "不适用"
                }

            policy_analysis["branch_policies"].append(branch_policy)

        return policy_analysis

    def calculate_tax_burden_comparison(self) -> Dict[str, Any]:
        """
        计算税负差异对比

        Returns:
            税负对比分析
        """
        if not self.headquarters:
            raise ValueError("未设置总机构信息")

        # 假设应纳税所得额（基于收入估算）
        estimated_profit = self.headquarters.total_income * 0.1  # 假设利润率10%

        # 不同情况下的税负
        standard_tax = estimated_profit * 0.25  # 法定25%

        if self.headquarters.is_high_tech:
            hq_tax = estimated_profit * 0.15  # 高新15%
        else:
            region = self.identify_region(self.headquarters.location)
            if region in ["西部地区", "海南自贸港"]:
                hq_tax = estimated_profit * 0.15  # 地区优惠15%
            else:
                hq_tax = standard_tax

        comparison = {
            "estimated_profit": estimated_profit,
            "scenarios": [
                {
                    "scenario": "一般企业（25%）",
                    "tax_rate": 25,
                    "tax_amount": standard_tax
                },
                {
                    "scenario": "高新技术企业（15%）",
                    "tax_rate": 15,
                    "tax_amount": estimated_profit * 0.15
                },
                {
                    "scenario": f"总部位于{self.identify_region(self.headquarters.location) or '一般地区'}（15%）",
                    "tax_rate": 15,
                    "tax_amount": hq_tax
                }
            ],
            "potential_saving": standard_tax - hq_tax
        }

        return comparison

    def run_full_analysis(self) -> Dict[str, Any]:
        """
        执行完整分析

        Returns:
            完整分析报告
        """
        allocation = self.calculate_allocation()
        policy_analysis = self.analyze_regional_policies()
        burden_comparison = self.calculate_tax_burden_comparison()

        report = {
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "headquarters": asdict(self.headquarters) if self.headquarters else None,
            "branches": [asdict(b) for b in self.branches],
            "total_tax": self.total_tax,
            "allocation_result": allocation,
            "policy_analysis": policy_analysis,
            "tax_burden_comparison": burden_comparison
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
    parser = argparse.ArgumentParser(description="跨区域企业所得税分摊计算脚本")
    parser.add_argument("--data-file", "-d", required=True, help="输入数据JSON文件路径")
    parser.add_argument("--output", "-o", help="输出报告JSON文件路径")

    args = parser.parse_args()

    # 创建计算器
    allocator = CrossRegionTaxAllocator()

    # 加载数据
    data = load_data_from_json(args.data_file)

    # 设置总机构
    allocator.set_headquarters(Headquarters(**data["headquarters"]))

    # 添加分支机构
    for item in data.get("branches", []):
        allocator.add_branch(Branch(**item))

    # 设置总税额
    allocator.set_total_tax(data["total_tax"])

    # 执行分析
    report = allocator.run_full_analysis()

    # 输出报告
    print("\n" + "="*60)
    print("跨区域经营企业所得税分摊分析报告")
    print("="*60)

    hq = report["headquarters"]
    print(f"\n总机构：{hq['name']}")
    print(f"所在地：{hq['location']}")
    print(f"是否高新技术企业：{'是' if hq['is_high_tech'] else '否'}")

    print(f"\n分支机构数量：{len(report['branches'])}")
    for branch in report["branches"]:
        print(f"  - {branch['name']}（{branch['location']}）")

    print(f"\n【总分摊计算】")
    alloc = report["allocation_result"]
    print(f"企业应纳所得税额：{alloc['total_tax']:,.2f} 元")
    print(f"总机构分摊（50%）：{alloc['hq_allocation']['tax_amount']:,.2f} 元")
    print(f"分支机构分摊（50%）：{alloc['total_branch_tax']:,.2f} 元")

    print(f"\n【分支机构分摊明细】")
    for ba in alloc["branch_allocations"]:
        print(f"  {ba['branch_name']}：{ba['pre_tax_amount']:,.2f} 元（分摊比例{ba['allocation_ratio']:.2f}%）")
        if ba["region"]:
            print(f"    所属区域：{ba['region']}")

    print(f"\n【政策分析】")
    pa = report["policy_analysis"]
    if "hq_policy" in pa:
        print(f"  总机构所在地政策：{pa['hq_policy']['policy']}")

    print(f"\n【税负对比】")
    tc = report["tax_burden_comparison"]
    print(f"  估算应纳税所得额：{tc['estimated_profit']:,.2f} 元")
    print(f"  潜在节税空间：{tc['potential_saving']:,.2f} 元")

    # 保存报告
    if args.output:
        save_report_to_json(report, args.output)
        print(f"\n报告已保存至：{args.output}")

    print("\n" + "="*60)


if __name__ == "__main__":
    main()
