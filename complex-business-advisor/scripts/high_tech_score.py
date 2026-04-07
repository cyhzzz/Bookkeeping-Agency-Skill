"""
高新企业认定自评打分

核心指标监控

输入：
- 研发费用数据
- 高新产品收入数据
- 科技人员数据
- 知识产权数据

输出：
- 六项核心指标达标情况
- 自评总分
- 不达标项及改进建议

评分标准（按高新企业认定管理办法）：
1. 知识产权（≤30分）
2. 科技成果转化能力（≤30分）
3. 研究开发组织管理水平（≤20分）
4. 企业成长性（≤20分）
5. 高新产品收入占比（≥60%得满分）
6. 研发费用占比（按销售收入分档）
"""

import json
import argparse
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class RDExpense:
    """研发费用明细"""
    year: int  # 年份
    rd_expense: float  # 研发费用


@dataclass
class HighTechProduct:
    """高新技术产品"""
    product_name: str  # 产品名称
    income: float  # 收入金额


@dataclass
class IPAsset:
    """知识产权"""
    ip_type: str  # 知识产权类型（发明专利/实用新型/软件著作权等）
    name: str  # 名称
    ownership: str  # 权属方式（自主研发/受让/并购等）
    year_obtained: int  # 取得年份


@dataclass
class SciTechPersonnel:
    """科技人员"""
    name: str  # 姓名
    position: str  # 岗位
    education: str  # 学历


class HighTechCertScorer:
    """高新企业认定评分器"""

    # 六项核心指标（一票否决项）
    CORE_CRITERIA = {
        "establishment_year": {
            "name": "成立时间",
            "requirement": "满1年",
            "passed": True  # 需根据企业实际情况判断
        },
        "high_income_ratio": {
            "name": "高新收入占比",
            "threshold": 60,  # ≥60%
            "passed": False
        },
        "rd_expense_ratio": {
            "name": "研发费用占比",
            "thresholds": [
                (50000000, 5),   # <5000万：≥5%
                (200000000, 4), # 5000万-2亿：≥4%
                (float('inf'), 3)  # >2亿：≥3%
            ],
            "passed": False
        },
        "sci_tech_personnel_ratio": {
            "name": "科技人员占比",
            "threshold": 10,  # ≥10%
            "passed": False
        },
        "cit_rate": {
            "name": "企业所得税率",
            "threshold": 15,
            "passed": False
        },
        "rd_management": {
            "name": "研发组织管理水平",
            "description": "通过评分考核",
            "passed": False
        }
    }

    # 评分指标权重
    SCORING_WEIGHTS = {
        "intellectual_property": 30,
        "tech_achievement_transformation": 30,
        "rd_organization_management": 20,
        "enterprise_growth": 20
    }

    # 知识产权评分标准
    IP_SCORING = {
        "advanced_degree": {
            "国际领先": 8,
            "国际先进": 6,
            "国内领先": 4,
            "国内先进": 2,
            "其他": 0
        },
        "core_support": {
            "强": 8,
            "较强": 6,
            "一般": 4,
            "弱": 2,
            "无": 0
        },
        "ip_count": {
            "1项及以上发明专利等Ⅰ类": 8,
            "5项及以上实用新型Ⅱ类": 6,
            "3-4项实用新型": 4,
            "1-3项": 2,
            "0项": 0
        },
        "ownership_mode": {
            "自主研发": 6,
            "受让受赠并购": 4,
            "独占许可": 3,
            "其他": 0
        }
    }

    # 科技成果转化评分
    TRANSFORMATION_SCORING = {
        "≥5项": (25, 30),
        "4项": (19, 24),
        "3项": (13, 18),
        "2项": (7, 12),
        "1项": (1, 6),
        "0项": (0, 0)
    }

    # 企业成长性评分
    GROWTH_SCORING = {
        "净资产增长率": {
            (35, float('inf')): (9, 10),
            (25, 35): (7, 8),
            (15, 25): (5, 6),
            (5, 15): (3, 4),
            (0, 5): (1, 2),
            (float('-inf'), 0): (0, 0)
        },
        "销售收入增长率": {
            (35, float('inf')): (9, 10),
            (25, 35): (7, 8),
            (15, 25): (5, 6),
            (5, 15): (3, 4),
            (0, 5): (1, 2),
            (float('-inf'), 0): (0, 0)
        }
    }

    def __init__(self):
        self.company_name: str = ""
        self.establishment_date: str = ""
        self.total_personnel: int = 0
        self.total_income: float = 0.0
        self.high_income: float = 0.0
        self.rd_expenses: List[RDExpense] = []
        self.high_tech_products: List[HighTechProduct] = []
        self.ip_assets: List[IPAsset] = []
        self.sci_tech_personnel: List[SciTechPersonnel] = []
        self.tech_transformation_count: int = 0
        self.cit_rate: float = 25.0  # 当前企业所得税率
        self.rd_management_score: float = 0.0  # 研发组织管理水平得分
        self.net_asset_growth: float = 0.0  # 净资产增长率
        self.sales_growth: float = 0.0  # 销售收入增长率

    def set_company_info(self, name: str, establishment_date: str,
                         total_personnel: int, total_income: float,
                         high_income: float, cit_rate: float):
        """设置企业基本信息"""
        self.company_name = name
        self.establishment_date = establishment_date
        self.total_personnel = total_personnel
        self.total_income = total_income
        self.high_income = high_income
        self.cit_rate = cit_rate

    def add_rd_expense(self, rd_expense: RDExpense):
        """添加研发费用数据"""
        self.rd_expenses.append(rd_expense)

    def add_ip_asset(self, ip_asset: IPAsset):
        """添加知识产权"""
        self.ip_assets.append(ip_asset)

    def add_sci_tech_personnel(self, personnel: SciTechPersonnel):
        """添加科技人员"""
        self.sci_tech_personnel.append(personnel)

    def check_establishment_year(self) -> Tuple[bool, str]:
        """检查成立时间"""
        try:
            est_year = int(self.establishment_date[:4])
            current_year = datetime.now().year
            years = current_year - est_year
            passed = years >= 1
            return passed, f"已成立{years}年（要求≥1年）"
        except:
            return False, "无法确定成立时间"

    def check_high_income_ratio(self) -> Tuple[bool, float, str]:
        """检查高新收入占比"""
        if self.total_income <= 0:
            return False, 0, "总收入为零或负数"
        ratio = (self.high_income / self.total_income) * 100
        passed = ratio >= 60
        return passed, ratio, f"占比{ratio:.1f}%（要求≥60%）"

    def check_rd_expense_ratio(self) -> Tuple[bool, float, str]:
        """检查研发费用占比"""
        if self.total_income <= 0:
            return False, 0, "总收入为零或负数"

        # 计算近三年研发费用合计和销售收入合计
        total_rd = sum(rd.rd_expense for rd in self.rd_expenses)
        total_income_sum = self.total_income * 3  # 简化计算

        if len(self.rd_expenses) >= 3:
            # 使用近三年平均
            avg_rd = total_rd / 3
            avg_income = total_income  # 用最近一年代替三年平均
        else:
            avg_rd = total_rd
            avg_income = self.total_income

        ratio = (avg_rd / avg_income) * 100 if avg_income > 0 else 0

        # 根据销售收入确定阈值
        threshold = 5  # 默认
        for income_threshold, req_ratio in self.CORE_CRITERIA["rd_expense_ratio"]["thresholds"]:
            if self.total_income < income_threshold:
                threshold = req_ratio
                break

        passed = ratio >= threshold
        return passed, ratio, f"占比{ratio:.2f}%（要求≥{threshold}%）"

    def check_personnel_ratio(self) -> Tuple[bool, float, str]:
        """检查科技人员占比"""
        if self.total_personnel <= 0:
            return False, 0, "总人数为零"
        ratio = (len(self.sci_tech_personnel) / self.total_personnel) * 100
        passed = ratio >= 10
        return passed, ratio, f"占比{ratio:.1f}%（要求≥10%）"

    def check_cit_rate(self) -> Tuple[bool, str]:
        """检查企业所得税率"""
        passed = self.cit_rate == 15
        return passed, f"税率{self.cit_rate}%（要求15%）"

    def score_intellectual_property(self) -> float:
        """
        评分：知识产权（≤30分）

        评分项：
        - 技术的先进程度（≤8）
        - 对主要产品的核心支持作用（≤8）
        - 知识产权数量（≤8）
        - 获取方式（≤6）
        """
        if not self.ip_assets:
            return 0.0

        # 分类统计知识产权
        invention_patents = [ip for ip in self.ip_assets if ip.ip_type == "发明专利"]
        utility_models = [ip for ip in self.ip_assets if ip.ip_type == "实用新型"]
        software_copyrights = [ip for ip in self.ip_assets if ip.ip_type == "软件著作权"]
        other_ips = [ip for ip in self.ip_assets if ip.ip_type not in ["发明专利", "实用新型", "软件著作权"]]

        # 1. 技术的先进程度（简化评估：假设企业自评）
        # 实际应按专家评审确定，此处使用默认分
        advanced_score = 4  # 国内一般水平

        # 2. 对主要产品的核心支持作用（简化评估）
        core_support_score = 6  # 较强

        # 3. 知识产权数量
        total_ips = len(self.ip_assets)
        type_1_count = len(invention_patents)  # Ⅰ类：发明专利
        type_2_count = len(utility_models) + len(other_ips)  # Ⅱ类：实用新型等

        if type_1_count >= 1:
            ip_count_score = 8
        elif type_2_count >= 5:
            ip_count_score = 6
        elif type_2_count >= 3:
            ip_count_score = 4
        elif type_2_count >= 1:
            ip_count_score = 2
        else:
            ip_count_score = 0

        # 4. 获取方式
        independent_count = len([ip for ip in self.ip_assets if ip.ownership == "自主研发"])
        transfer_count = len([ip for ip in self.ip_assets if ip.ownership in ["受让", "并购", "受赠"]])

        if independent_count == len(self.ip_assets):
            ownership_score = 6
        elif transfer_count > 0:
            ownership_score = 4
        else:
            ownership_score = 3

        total = advanced_score + core_support_score + ip_count_score + ownership_score
        return min(total, 30)

    def score_tech_transformation(self) -> float:
        """
        评分：科技成果转化能力（≤30分）
        """
        count = self.tech_transformation_count

        if count >= 5:
            return 27.5  # 取中间值
        elif count == 4:
            return 21.5
        elif count == 3:
            return 15.5
        elif count == 2:
            return 9.5
        elif count == 1:
            return 3.5
        else:
            return 0

    def score_rd_management(self) -> float:
        """
        评分：研发组织管理水平（≤20分）

        评分项：
        - 研发组织管理制度（≤6）
        - 研发机构设置（≤6）
        - 成果转化激励制度（≤4）
        - 人员培训制度（≤2）
        - 开放式创新平台（≤2）
        """
        return min(self.rd_management_score, 20)

    def score_enterprise_growth(self) -> float:
        """
        评分：企业成长性（≤20分）

        - 净资产增长率（≤10）
        - 销售收入增长率（≤10）
        """
        def get_growth_score(growth_rate: float) -> float:
            if growth_rate >= 35:
                return 9.5
            elif growth_rate >= 25:
                return 7.5
            elif growth_rate >= 15:
                return 5.5
            elif growth_rate >= 5:
                return 3.5
            elif growth_rate > 0:
                return 1.5
            else:
                return 0

        net_asset_score = get_growth_score(self.net_asset_growth)
        sales_score = get_growth_score(self.sales_growth)

        return min(net_asset_score + sales_score, 20)

    def run_full_assessment(self) -> Dict[str, Any]:
        """
        执行完整评估

        Returns:
            评估报告
        """
        # 检查六项核心指标
        core_results = {}

        passed, detail = self.check_establishment_year()
        core_results["establishment_year"] = {
            "passed": passed,
            "detail": detail
        }

        passed, ratio, detail = self.check_high_income_ratio()
        core_results["high_income_ratio"] = {
            "passed": passed,
            "value": ratio,
            "detail": detail
        }

        passed, ratio, detail = self.check_rd_expense_ratio()
        core_results["rd_expense_ratio"] = {
            "passed": passed,
            "value": ratio,
            "detail": detail
        }

        passed, ratio, detail = self.check_personnel_ratio()
        core_results["sci_tech_personnel_ratio"] = {
            "passed": passed,
            "value": ratio,
            "detail": detail
        }

        passed, detail = self.check_cit_rate()
        core_results["cit_rate"] = {
            "passed": passed,
            "detail": detail
        }

        core_results["rd_management"] = {
            "passed": self.rd_management_score >= 12,  # 假设12分为及格线
            "detail": f"研发管理水平评分{self.rd_management_score}分"
        }

        # 统计核心指标通过数
        core_pass_count = sum(1 for v in core_results.values() if v["passed"])

        # 评分指标计算
        ip_score = self.score_intellectual_property()
        transformation_score = self.score_tech_transformation()
        management_score = self.score_rd_management()
        growth_score = self.score_enterprise_growth()
        total_score = ip_score + transformation_score + management_score + growth_score

        # 判断是否达标
        all_core_passed = all(v["passed"] for v in core_results.values())
        score_passed = total_score >= 71
        overall_passed = all_core_passed and score_passed

        # 生成改进建议
        improvements = []

        if not core_results["high_income_ratio"]["passed"]:
            improvements.append({
                "item": "高新收入占比",
                "current": core_results["high_income_ratio"]["value"],
                "target": 60,
                "suggestion": "加大高新技术产品研发投入，提升高新收入占比"
            })

        if not core_results["rd_expense_ratio"]["passed"]:
            improvements.append({
                "item": "研发费用占比",
                "current": core_results["rd_expense_ratio"]["value"],
                "target": "按收入档位",
                "suggestion": "增加研发投入，确保研发费用符合销售收入对应比例"
            })

        if not core_results["sci_tech_personnel_ratio"]["passed"]:
            improvements.append({
                "item": "科技人员占比",
                "current": core_results["sci_tech_personnel_ratio"]["value"],
                "target": 10,
                "suggestion": "招聘或转化更多科技人员"
            })

        if ip_score < 20:
            improvements.append({
                "item": "知识产权得分",
                "current": ip_score,
                "target": 20,
                "suggestion": "加强自主研发，增加发明专利申请"
            })

        if transformation_score < 20:
            improvements.append({
                "item": "科技成果转化",
                "current": transformation_score,
                "target": 20,
                "suggestion": "建立科技成果转化机制，确保每年转化4项以上"
            })

        if management_score < 12:
            improvements.append({
                "item": "研发组织管理",
                "current": management_score,
                "target": 12,
                "suggestion": "完善研发管理制度，建立专门研发机构"
            })

        report = {
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "company_name": self.company_name,
            "core_criteria_results": core_results,
            "core_pass_count": core_pass_count,
            "total_core_criteria": 6,
            "scoring_results": {
                "intellectual_property": {
                    "score": ip_score,
                    "max": 30,
                    "items": {
                        "advanced_degree": "技术先进程度",
                        "core_support": "核心支持作用",
                        "ip_count": "知识产权数量",
                        "ownership_mode": "获取方式"
                    }
                },
                "tech_transformation": {
                    "score": transformation_score,
                    "max": 30,
                    "transformation_count": self.tech_transformation_count
                },
                "rd_management": {
                    "score": management_score,
                    "max": 20
                },
                "enterprise_growth": {
                    "score": growth_score,
                    "max": 20,
                    "net_asset_growth": self.net_asset_growth,
                    "sales_growth": self.sales_growth
                }
            },
            "total_score": total_score,
            "pass_threshold": 71,
            "overall_passed": overall_passed,
            "assessment_result": {
                "core_indicators": "通过" if all_core_passed else "未通过",
                "scoring": "通过" if score_passed else "未通过",
                "overall": "通过" if overall_passed else "待改进"
            },
            "improvements": improvements
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
    parser = argparse.ArgumentParser(description="高新企业认定自评打分脚本")
    parser.add_argument("--data-file", "-d", required=True, help="输入数据JSON文件路径")
    parser.add_argument("--output", "-o", help="输出报告JSON文件路径")

    args = parser.parse_args()

    # 创建评分器
    scorer = HighTechCertScorer()

    # 加载数据
    data = load_data_from_json(args.data_file)

    # 设置企业信息
    scorer.set_company_info(
        name=data["company_name"],
        establishment_date=data["establishment_date"],
        total_personnel=data["total_personnel"],
        total_income=data["total_income"],
        high_income=data["high_income"],
        cit_rate=data.get("cit_rate", 25)
    )

    # 添加研发费用
    for item in data.get("rd_expenses", []):
        scorer.add_rd_expense(RDExpense(**item))

    # 添加知识产权
    for item in data.get("ip_assets", []):
        scorer.add_ip_asset(IPAsset(**item))

    # 添加科技人员
    for item in data.get("sci_tech_personnel", []):
        scorer.add_sci_tech_personnel(SciTechPersonnel(**item))

    # 设置其他评分参数
    scorer.tech_transformation_count = data.get("tech_transformation_count", 0)
    scorer.rd_management_score = data.get("rd_management_score", 0)
    scorer.net_asset_growth = data.get("net_asset_growth", 0)
    scorer.sales_growth = data.get("sales_growth", 0)

    # 执行评估
    report = scorer.run_full_assessment()

    # 输出报告
    print("\n" + "="*60)
    print(f"高新企业认定自评报告")
    print("="*60)
    print(f"\n企业名称：{report['company_name']}")

    print(f"\n【六项核心指标】({report['core_pass_count']}/6 通过)")
    for key, result in report["core_criteria_results"].items():
        status = "✓" if result["passed"] else "✗"
        detail = result.get("detail", "")
        print(f"  {status} {detail}")

    print(f"\n【评分指标】")
    scoring = report["scoring_results"]
    print(f"  知识产权：{scoring['intellectual_property']['score']}/{scoring['intellectual_property']['max']}分")
    print(f"  科技成果转化：{scoring['tech_transformation']['score']}/{scoring['tech_transformation']['max']}分")
    print(f"  研发组织管理：{scoring['rd_management']['score']}/{scoring['rd_management']['max']}分")
    print(f"  企业成长性：{scoring['enterprise_growth']['score']}/{scoring['enterprise_growth']['max']}分")
    print(f"  总分：{report['total_score']}分（要求≥71分）")

    print(f"\n【评审结论】")
    result = report["assessment_result"]
    print(f"  核心指标：{result['core_indicators']}")
    print(f"  评分项：{result['scoring']}")
    print(f"  综合判定：{result['overall']}")

    if report["improvements"]:
        print(f"\n【改进建议】({len(report['improvements'])}项)")
        for imp in report["improvements"]:
            print(f"  - {imp['item']}：{imp['suggestion']}")

    # 保存报告
    if args.output:
        save_report_to_json(report, args.output)
        print(f"\n报告已保存至：{args.output}")

    print("\n" + "="*60)


if __name__ == "__main__":
    main()
