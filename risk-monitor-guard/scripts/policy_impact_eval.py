"""
政策变动影响评估脚本

输入：
- 新政策文本
- 客户列表

输出：
- 受影响客户清单
- 影响程度评估
- 操作建议

评估逻辑：
1. 政策适用性判断
   - 企业类型匹配（一般纳税人/小规模/高新等）
   - 行业匹配
   - 地区匹配

2. 影响金额测算
   - 税负变化
   - 申报要求变化

3. 过渡期安排
   - 生效时间
   - 优惠享受条件
"""

import json
import re
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime, date


@dataclass
class PolicyInfo:
    """政策信息"""
    policy_id: str
    title: str
    source: str
    issued_date: str
    effective_date: str
    category: str
    summary: str
    impact_level: str  # high/medium/low
    key_changes: List[str]
    transition_period: Optional[str] = None
    conditions: Optional[List[str]] = None


@dataclass
class CustomerImpact:
    """客户影响评估"""
    customer_id: str
    customer_name: str
    impact_type: str  # positive/negative/neutral
    impact_level: str  # high/medium/low
    estimated_amount: float  # 预估影响金额（元）
    reason: str
    action_required: List[str]


@dataclass
class PolicyEvaluation:
    """政策评估结果"""
    policy: PolicyInfo
    affected_customers: List[CustomerImpact]
    total_positive_count: int
    total_negative_count: int
    total_neutral_count: int
    estimated_total_benefit: float
    estimated_total_cost: float
    evaluation_time: str
    suggestions: List[str]


class PolicyImpactEvaluator:
    """政策影响评估器"""

    # 政策类别与客户类型映射
    POLICY_CUSTOMER_TYPE_MAP = {
        "small_scale_benefit": ["小规模纳税人"],
        "general_taxpayer_benefit": ["一般纳税人"],
        "high_tech_benefit": ["高新技术企业", "软件企业"],
        "micro_enterprise_benefit": ["小型微利企业"],
        "rd_expense_benefit": ["高新技术企业", "科技型中小企业"],
        "invoice_reform": ["一般纳税人"],
        "tax_withholding": ["扣缴义务人"],
        "fapiao_management": ["一般纳税人", "小规模纳税人"],
    }

    # 政策类别与行业映射
    POLICY_INDUSTRY_MAP = {
        "tech_industry_benefit": ["软件开发", "信息技术服务"],
        "manufacturing_benefit": ["制造业"],
        "service_industry_benefit": ["服务业"],
        "environmental_protection": ["环保相关"],
        "agricultural_benefit": ["农业"],
    }

    def __init__(self):
        self.evaluation_counter = 0

    def _generate_policy_id(self) -> str:
        self.evaluation_counter += 1
        return f"POL-EVAL-{datetime.now().strftime('%Y%m%d')}-{self.evaluation_counter:03d}"

    def parse_policy(self, policy_text: str) -> PolicyInfo:
        """
        从政策文本中解析关键信息

        参数：
            policy_text: 政策文件文本

        返回：
            PolicyInfo对象
        """
        # 简化解析逻辑，实际应用中需要结合LLM或规则引擎
        lines = policy_text.strip().split('\n')

        info = {
            "policy_id": self._generate_policy_id(),
            "title": "",
            "source": "",
            "issued_date": "",
            "effective_date": "",
            "category": "",
            "summary": "",
            "impact_level": "medium",
            "key_changes": [],
        }

        current_key = None
        content_buffer = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 匹配标题
            if line.startswith("#") or line.startswith("【"):
                continue

            # 匹配关键字段
            if "标题" in line or "名称" in line:
                current_key = "title"
                info["title"] = line.split("：")[-1].split("】")[-1].strip()
                continue
            if "来源" in line or "发文机关" in line:
                current_key = "source"
                info["source"] = line.split("：")[-1].strip()
                continue
            if "发文日期" in line or "发布日期" in line:
                current_key = "issued_date"
                info["issued_date"] = line.split("：")[-1].strip()
                continue
            if "生效日期" in line or "实施日期" in line:
                current_key = "effective_date"
                info["effective_date"] = line.split("：")[-1].strip()
                continue
            if "类别" in line:
                current_key = "category"
                info["category"] = line.split("：")[-1].strip()
                continue
            if "摘要" in line or "变化摘要" in line:
                current_key = "summary"
                info["summary"] = line.split("：")[-1].strip()
                continue
            if "主要变化" in line or "变化点" in line:
                current_key = "key_changes"
                info["key_changes"].append(line.split("：")[-1].strip())
                continue

        # 简化返回
        return PolicyInfo(
            policy_id=info["policy_id"],
            title=info["title"] or "未识别政策标题",
            source=info["source"] or "未知来源",
            issued_date=info["issued_date"] or datetime.now().strftime("%Y-%m-%d"),
            effective_date=info["effective_date"] or datetime.now().strftime("%Y-%m-%d"),
            category=info["category"] or "其他",
            summary=info["summary"] or "政策内容摘要",
            impact_level=info["impact_level"],
            key_changes=info["key_changes"],
            transition_period=None,
            conditions=None
        )

    def assess_customer_eligibility(
        self,
        customer: Dict,
        policy: PolicyInfo
    ) -> Tuple[bool, str]:
        """
        评估客户是否符合政策适用条件

        参数：
            customer: 客户信息
                - customer_id: 客户编号
                - customer_name: 客户名称
                - tax_type: 纳税人类型（一般纳税人/小规模纳税人）
                - industry: 所属行业
                - is_high_tech: 是否高新企业
                - is_small_micro: 是否小微
                - annual_revenue: 年营业收入
                - annual_tax: 年纳税额
            policy: 政策信息

        返回：
            (是否适用, 原因说明)
        """
        category = policy.category

        # 检查纳税人类型
        tax_type = customer.get("tax_type", "")
        if "小微" in category or "小规模" in category:
            if tax_type != "小规模纳税人":
                return False, "该政策仅适用于小规模纳税人"
        if "一般纳税人" in category or "增值税" in category:
            if tax_type not in ["一般纳税人"]:
                return False, "该政策主要影响一般纳税人"

        # 检查企业资质
        if "高新" in category or "研发" in category:
            if not customer.get("is_high_tech", False):
                return False, "该政策仅适用于高新技术企业"

        # 检查小微条件
        if "小微" in category:
            if not customer.get("is_small_micro", False):
                annual_revenue = customer.get("annual_revenue", 0)
                if annual_revenue > 5000000:  # 小微标准
                    return False, "企业规模超出小微企业标准"

        # 检查行业
        industry = customer.get("industry", "")
        if "制造业" in category and "制造" not in industry:
            return False, "该政策仅适用于制造业"

        return True, "符合政策适用条件"

    def calc_impact_amount(
        self,
        customer: Dict,
        policy: PolicyInfo
    ) -> Tuple[str, float, str]:
        """
        计算政策对客户的预估影响金额

        参数：
            customer: 客户信息
            policy: 政策信息

        返回：
            (影响类型, 预估金额, 说明)
        """
        annual_tax = customer.get("annual_tax", 0)
        annual_revenue = customer.get("annual_revenue", 0)

        impact_type = "neutral"
        amount = 0.0
        reason = "该政策对贵公司无显著影响"

        category = policy.category.lower()

        # 税收优惠政策影响测算
        if "优惠" in category or "减免" in category:
            if "增值" in category:
                # 增值税优惠
                if annual_revenue > 0:
                    # 假设计算税率优惠
                    benefit_rate = 0.01  # 假设优惠1%
                    amount = annual_revenue * benefit_rate
                    impact_type = "positive"
                    reason = f"预计年减税约{amount:.0f}元（按收入{annual_revenue}万元计算）"
            elif "所得" in category:
                # 企业所得税优惠
                if annual_tax > 0:
                    benefit_rate = 0.05  # 假设优惠5%
                    amount = annual_tax * benefit_rate
                    impact_type = "positive"
                    reason = f"预计年减税约{amount:.0f}元（按年所得额计算）"

        # 税率调整影响测算
        elif "税率" in category or "调整" in category:
            if annual_tax > 0:
                change_rate = 0.005  # 假设税率变化0.5%
                amount = annual_revenue * change_rate
                impact_type = "negative"
                reason = f"预计年增税约{amount:.0f}元（税负有所上升）"

        # 申报要求变化（主要影响人力成本）
        elif "申报" in category or "报送" in category:
            impact_type = "neutral"
            reason = "申报要求变化主要影响操作流程，预计增加一定工作量"

        return impact_type, amount, reason

    def generate_action_items(
        self,
        policy: PolicyInfo,
        customer: Dict,
        impact_type: str
    ) -> List[str]:
        """
        生成客户需要执行的操作清单

        参数：
            policy: 政策信息
            customer: 客户信息
            impact_type: 影响类型

        返回：
            操作建议列表
        """
        actions = []
        category = policy.category

        # 基于政策类型生成操作建议
        if "优惠" in category:
            actions.append("确认是否仍符合优惠享受条件")
            if policy.effective_date:
                actions.append(f"关注政策生效日期：{policy.effective_date}")
            actions.append("准备优惠备案相关材料（如需）")
            actions.append("更新所得税预缴计算逻辑")

        if "增值税" in category:
            actions.append("检查进项发票取得情况")
            actions.append("确认销项税额计算是否正确")
            actions.append("更新发票认证截止日期管理")

        if "申报" in category:
            actions.append("更新申报日历设置")
            actions.append("准备好新申报表格式")
            actions.append("关注申报截止日期变化")

        if "发票" in category:
            actions.append("检查现有发票使用情况")
            actions.append("关注数电发票推广进度")
            actions.append("完成开票系统升级对接")

        if impact_type == "positive":
            actions.append("可在电子税务局提交优惠备案申请")
        elif impact_type == "negative":
            actions.append("提前做好税负上升的资金安排")
            actions.append("与财务顾问讨论成本控制方案")

        return actions

    def evaluate_policy_impact(
        self,
        policy: PolicyInfo,
        customers: List[Dict]
    ) -> PolicyEvaluation:
        """
        评估政策对客户群体的影响

        参数：
            policy: 政策信息
            customers: 客户列表

        返回：
            PolicyEvaluation对象
        """
        affected_customers = []
        total_positive = 0
        total_negative = 0
        total_neutral = 0
        total_benefit = 0.0
        total_cost = 0.0

        for customer in customers:
            # 评估适用性
            is_eligible, reason = self.assess_customer_eligibility(customer, policy)

            if not is_eligible:
                continue

            # 计算影响金额
            impact_type, amount, calc_reason = self.calc_impact_amount(customer, policy)

            # 生成操作清单
            actions = self.generate_action_items(policy, customer, impact_type)

            # 确定影响等级
            if impact_type == "positive":
                impact_level = "high" if amount > 50000 else "medium" if amount > 10000 else "low"
                total_positive += 1
                total_benefit += amount
            elif impact_type == "negative":
                impact_level = "high" if amount > 50000 else "medium" if amount > 10000 else "low"
                total_negative += 1
                total_cost += amount
            else:
                impact_level = "low"
                total_neutral += 1

            affected_customers.append(CustomerImpact(
                customer_id=customer.get("customer_id", ""),
                customer_name=customer.get("customer_name", ""),
                impact_type=impact_type,
                impact_level=impact_level,
                estimated_amount=amount,
                reason=f"{calc_reason}；{reason}",
                action_required=actions
            ))

        # 生成整体建议
        suggestions = self._generate_overall_suggestions(policy, affected_customers)

        return PolicyEvaluation(
            policy=policy,
            affected_customers=affected_customers,
            total_positive_count=total_positive,
            total_negative_count=total_negative,
            total_neutral_count=total_neutral,
            estimated_total_benefit=total_benefit,
            estimated_total_cost=total_cost,
            evaluation_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            suggestions=suggestions
        )

    def _generate_overall_suggestions(
        self,
        policy: PolicyInfo,
        affected_customers: List[CustomerImpact]
    ) -> List[str]:
        """生成整体建议"""
        suggestions = []

        # 按影响类型分类
        positive_customers = [c for c in affected_customers if c.impact_type == "positive"]
        negative_customers = [c for c in affected_customers if c.impact_type == "negative"]

        if positive_customers:
            high_impact = [c for c in positive_customers if c.impact_level == "high"]
            if high_impact:
                suggestions.append(
                    f"有{len(high_impact)}户高影响客户预计享受较大税收优惠，建议优先跟进办理备案"
                )

        if negative_customers:
            high_impact = [c for c in negative_customers if c.impact_level == "high"]
            if high_impact:
                suggestions.append(
                    f"有{len(high_impact)}户客户将面临较大税负上升，建议提前做好沟通和筹划"
                )

        # 政策层面建议
        if policy.transition_period:
            suggestions.append(f"关注过渡期安排：{policy.transition_period}")

        if policy.key_changes:
            suggestions.append(f"政策重点变化：{'、'.join(policy.key_changes[:3])}")

        suggestions.append("建议组织内部培训，确保相关人员了解政策变化内容")
        suggestions.append("及时在系统中更新政策信息，确保申报口径准确")

        return suggestions

    def generate_bulletin(
        self,
        evaluation: PolicyEvaluation
    ) -> str:
        """
        生成政策变动公告文本

        参数：
            evaluation: 政策评估结果

        返回：
            公告文本（Markdown格式）
        """
        policy = evaluation.policy

        lines = [
            f"# {policy.title}",
            "",
            f"**发布机构**：{policy.source}",
            f"**发布日期**：{policy.issued_date}",
            f"**生效日期**：{policy.effective_date}",
            f"**政策类别**：{policy.category}",
            "",
            "---",
            "",
            "## 一、政策概述",
            "",
            policy.summary,
            "",
        ]

        if policy.key_changes:
            lines.append("## 二、主要变化点")
            lines.append("")
            for i, change in enumerate(policy.key_changes, 1):
                lines.append(f"{i}. {change}")
            lines.append("")

        lines.append("## 三、影响分析")
        lines.append("")
        lines.append(f"| 影响类型 | 客户数量 |")
        lines.append(f"|----------|----------|")
        lines.append(f"| 正面影响 | {evaluation.total_positive_count}户 |")
        lines.append(f"| 负面影响 | {evaluation.total_negative_count}户 |")
        lines.append(f"| 中性影响 | {evaluation.total_neutral_count}户 |")
        lines.append("")
        lines.append(f"**预估总减税**：{evaluation.estimated_total_benefit:,.0f}元")
        lines.append(f"**预估总增税**：{evaluation.estimated_total_cost:,.0f}元")
        lines.append("")

        if evaluation.affected_customers:
            lines.append("## 四、受影响客户清单")
            lines.append("")
            lines.append("| 客户名称 | 影响类型 | 预估金额 | 主要措施 |")
            lines.append("|----------|----------|----------|----------|")
            for c in evaluation.affected_customers[:20]:  # 最多显示20户
                impact_symbol = "+" if c.impact_type == "positive" else "-" if c.impact_type == "negative" else "~"
                actions = "、".join(c.action_required[:2])
                lines.append(f"| {c.customer_name} | {impact_symbol} | {c.estimated_amount:,.0f}元 | {actions} |")
            lines.append("")

        if evaluation.suggestions:
            lines.append("## 五、操作建议")
            lines.append("")
            for i, s in enumerate(evaluation.suggestions, 1):
                lines.append(f"{i}. {s}")
            lines.append("")

        lines.append("---")
        lines.append(f"*评估时间：{evaluation.evaluation_time}*")

        return "\n".join(lines)


def demo():
    """演示用法"""
    evaluator = PolicyImpactEvaluator()

    # 示例政策
    policy_text = """
    标题：关于小微企业和个体工商户所得税优惠政策的公告
    来源：国家税务总局 财政部
    发文日期：2024-01-01
    生效日期：2024-01-01
    类别：小微企业所得税优惠
    摘要：对小型微利企业年应纳税所得额不超过300万元的部分，减按25%计入应纳税所得额，按20%的税率缴纳企业所得税
    主要变化：
    - 优惠力度调整
    - 适用范围扩大
    """

    policy = evaluator.parse_policy(policy_text)
    print(f"解析政策ID: {policy.policy_id}")
    print(f"政策标题: {policy.title}")

    # 示例客户列表
    customers = [
        {
            "customer_id": "C001",
            "customer_name": "上海XX商贸有限公司",
            "tax_type": "小规模纳税人",
            "industry": "百货零售",
            "is_high_tech": False,
            "is_small_micro": True,
            "annual_revenue": 3000000,
            "annual_tax": 80000,
        },
        {
            "customer_id": "C002",
            "customer_name": "北京YY科技有限公司",
            "tax_type": "一般纳税人",
            "industry": "软件开发",
            "is_high_tech": True,
            "is_small_micro": False,
            "annual_revenue": 50000000,
            "annual_tax": 2000000,
        },
        {
            "customer_id": "C003",
            "customer_name": "深圳ZZ制造有限公司",
            "tax_type": "一般纳税人",
            "industry": "电子设备制造",
            "is_high_tech": False,
            "is_small_micro": False,
            "annual_revenue": 80000000,
            "annual_tax": 3500000,
        },
    ]

    # 评估影响
    evaluation = evaluator.evaluate_policy_impact(policy, customers)

    print("\n" + "=" * 60)
    print(f"政策影响评估报告")
    print("=" * 60)
    print(f"政策：{policy.title}")
    print(f"评估时间：{evaluation.evaluation_time}")
    print("-" * 60)
    print(f"受影响客户总数：{len(evaluation.affected_customers)}户")
    print(f"  - 正面影响：{evaluation.total_positive_count}户")
    print(f"  - 负面影响：{evaluation.total_negative_count}户")
    print(f"  - 中性影响：{evaluation.total_neutral_count}户")
    print(f"预估总减税：{evaluation.estimated_total_benefit:,.0f}元")
    print(f"预估总增税：{evaluation.estimated_total_cost:,.0f}元")
    print("-" * 60)
    print("受影响客户：")
    for c in evaluation.affected_customers:
        symbol = "+" if c.impact_type == "positive" else "-" if c.impact_type == "negative" else "~"
        print(f"  [{symbol}] {c.customer_name}: {c.reason}")
    print("-" * 60)
    print("操作建议：")
    for i, s in enumerate(evaluation.suggestions, 1):
        print(f"  {i}. {s}")
    print("=" * 60)

    # 生成公告
    print("\n生成公告文本：")
    bulletin = evaluator.generate_bulletin(evaluation)
    print(bulletin)


if __name__ == "__main__":
    demo()
