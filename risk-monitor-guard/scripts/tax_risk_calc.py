"""
税务风险指标计算脚本
多维度风险评估

输入：
- 申报数据（增值税申报表、企业所得税申报表等）
- 发票数据（进项、销项）
- 行业基准

输出：
- 风险指标得分
- 风险等级（红/橙/黄/绿）
- 详细风险清单

风险指标：
1. 税负率偏离度
   - 异常标准：偏离行业均值超过30%
   - 得分：正常=0，低风险=10，中风险=30，高风险=50

2. 发票合规性
   - 作废发票占比（正常<5%）
   - 红字发票频率
   - 异常金额发票

3. 申报及时性
   - 逾期次数
   - 逾期天数

4. 优惠备案有效性
   - 备案过期
   - 备案与实际不符
"""

import json
import math
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime, date


@dataclass
class RiskIndicator:
    """风险指标"""
    type: str
    value: float
    benchmark: float
    deviation: float
    score: int
    level: str
    description: str


@dataclass
class RiskAlert:
    """风险预警"""
    risk_id: str
    customer_id: str
    risk_type: str
    risk_level: str
    risk_description: str
    basis: str
    suggestion: str
    detected_at: str
    status: str = "pending"


class TaxRiskCalculator:
    """税务风险计算器"""

    # 行业税负率基准（增值税）
    INDUSTRY_VAT_BENCHMARKS = {
        # 制造业
        "通用设备制造": 0.038,
        "专用设备制造": 0.042,
        "电子设备制造": 0.035,
        "纺织服装制造": 0.038,
        "食品饮料制造": 0.045,
        "医药制造": 0.050,
        "汽车制造": 0.032,
        "化工原料制造": 0.040,
        # 服务业
        "餐饮服务": 0.032,
        "酒店住宿": 0.038,
        "咨询服务": 0.055,
        "软件开发": 0.040,
        "广告服务": 0.050,
        "代理记账": 0.045,
        "人力资源服务": 0.038,
        # 建筑业
        "房屋建筑": 0.032,
        "土木工程": 0.035,
        "建筑安装": 0.038,
        "装饰装修": 0.042,
        # 批发零售业
        "百货零售": 0.022,
        "超市零售": 0.018,
        "医药零售": 0.032,
        "农资零售": 0.028,
        "建材批发": 0.028,
        "电子产品批发": 0.032,
        # 交通运输业
        "公路货运": 0.032,
        "水路货运": 0.028,
        "航空运输": 0.038,
        "仓储物流": 0.038,
        # 房地产开发
        "房地产开发": 0.045,
        "物业管理": 0.038,
        # 金融业
        "银行": 0.050,
        "保险": 0.045,
        "证券": 0.055,
        # 默认值
        "default": 0.035,
    }

    # 行业企业所得税税负率基准
    INDUSTRY_EIT_BENCHMARKS = {
        "通用设备制造": 0.028,
        "专用设备制造": 0.032,
        "电子设备制造": 0.025,
        "纺织服装制造": 0.022,
        "食品饮料制造": 0.030,
        "医药制造": 0.035,
        "汽车制造": 0.022,
        "化工原料制造": 0.028,
        "餐饮服务": 0.022,
        "酒店住宿": 0.025,
        "咨询服务": 0.035,
        "软件开发": 0.028,
        "广告服务": 0.030,
        "代理记账": 0.028,
        "人力资源服务": 0.025,
        "房屋建筑": 0.022,
        "土木工程": 0.025,
        "建筑安装": 0.022,
        "装饰装修": 0.028,
        "百货零售": 0.013,
        "超市零售": 0.010,
        "医药零售": 0.018,
        "农资零售": 0.015,
        "建材批发": 0.016,
        "电子产品批发": 0.018,
        "公路货运": 0.018,
        "水路货运": 0.016,
        "航空运输": 0.022,
        "仓储物流": 0.022,
        "房地产开发": 0.030,
        "物业管理": 0.022,
        "银行": 0.035,
        "保险": 0.030,
        "证券": 0.038,
        "default": 0.025,
    }

    def __init__(self):
        self.risk_counter = 0

    def _get_benchmark(self, industry: str, tax_type: str = "vat") -> float:
        """获取行业基准税负率"""
        benchmarks = self.INDUSTRY_VAT_BENCHMARKS if tax_type == "vat" else self.INDUSTRY_EIT_BENCHMARKS
        return benchmarks.get(industry, benchmarks["default"])

    def _calc_deviation(self, actual: float, expected: float) -> float:
        """计算偏离度"""
        if expected == 0:
            return 0.0
        return abs(actual - expected) / expected

    def _deviation_to_score(self, deviation: float) -> Tuple[int, str]:
        """偏离度转评分和等级"""
        if deviation < 0.15:
            return 0, "green"
        elif deviation < 0.30:
            return 10, "yellow"
        elif deviation < 0.50:
            return 30, "orange"
        else:
            return 50, "red"

    def _generate_risk_id(self) -> str:
        """生成风险ID"""
        self.risk_counter += 1
        return f"TR-{datetime.now().strftime('%Y%m%d')}-{self.risk_counter:03d}"

    def calc_burden_deviation(
        self,
        tax_amount: float,
        sales: float,
        industry: str,
        tax_type: str = "vat"
    ) -> RiskIndicator:
        """
        计算税负率偏离度

        参数：
            tax_amount: 实缴税额
            sales: 销售收入
            industry: 行业
            tax_type: 税种类型（vat/eit）

        返回：
            RiskIndicator对象
        """
        expected_rate = self._get_benchmark(industry, tax_type)
        actual_rate = tax_amount / sales if sales > 0 else 0.0
        deviation = self._calc_deviation(actual_rate, expected_rate)
        score, level = self._deviation_to_score(deviation)

        if actual_rate < expected_rate:
            direction = "偏低"
        else:
            direction = "偏高"

        description = (
            f"{'增值税' if tax_type == 'vat' else '企业所得税'}税负率{direction}："
            f"实际{actual_rate*100:.2f}% vs 行业均值{expected_rate*100:.2f}%，"
            f"偏离{deviation*100:.1f}%"
        )

        return RiskIndicator(
            type=f"tax_burden_{tax_type}",
            value=actual_rate,
            benchmark=expected_rate,
            deviation=deviation,
            score=score,
            level=level,
            description=description
        )

    def calc_invoice_compliance(
        self,
        voided_count: int,
        total_count: int,
        red_invoice_months: int,
        abnormal_amount_count: int
    ) -> RiskIndicator:
        """
        计算发票合规性风险

        参数：
            voided_count: 作废发票份数
            total_count: 开具发票总份数
            red_invoice_months: 连续出现红字发票月数
            abnormal_amount_count: 异常金额发票张数

        返回：
            RiskIndicator对象
        """
        voided_ratio = voided_count / total_count if total_count > 0 else 0.0

        # 计算各子项得分
        voided_score = 0
        red_score = 0
        abnormal_score = 0

        # 作废发票占比（>5%开始扣分）
        if voided_ratio > 0.05:
            voided_score = min(15, int((voided_ratio - 0.05) * 300))

        # 红字发票（连续3个月开始扣分）
        if red_invoice_months >= 3:
            red_score = min(20, red_invoice_months * 5)

        # 异常金额发票（每张5分）
        abnormal_score = min(15, abnormal_amount_count * 5)

        total_score = voided_score + red_score + abnormal_score

        # 确定等级
        if total_score >= 40:
            level = "red"
        elif total_score >= 20:
            level = "orange"
        elif total_score >= 10:
            level = "yellow"
        else:
            level = "green"

        details = []
        if voided_ratio > 0.05:
            details.append(f"作废发票占比{voided_ratio*100:.1f}%超过5%阈值")
        if red_invoice_months >= 3:
            details.append(f"连续{red_invoice_months}个月出现红字发票")
        if abnormal_amount_count > 0:
            details.append(f"发现{abnormal_amount_count}张异常金额发票")

        description = "；".join(details) if details else "发票合规性正常"

        return RiskIndicator(
            type="invoice_compliance",
            value=voided_ratio,
            benchmark=0.05,
            deviation=voided_ratio,
            score=total_score,
            level=level,
            description=description
        )

    def calc_filing_timeliness(
        self,
        overdue_months: int,
        overdue_days: int
    ) -> RiskIndicator:
        """
        计算申报及时性风险

        参数：
            overdue_months: 连续逾期月数
            overdue_days: 累计逾期天数

        返回：
            RiskIndicator对象
        """
        # 根据连续逾期月数确定等级
        if overdue_months >= 3:
            score = 50
            level = "red"
        elif overdue_months == 2:
            score = 30
            level = "orange"
        elif overdue_months == 1:
            score = 15
            level = "yellow"
        else:
            score = 0
            level = "green"

        # 额外逾期天数加严
        if overdue_days > 30:
            score = min(50, score + 10)
            level = "red" if level != "red" else level
        elif overdue_days > 10:
            if level == "green":
                score = 10
                level = "yellow"

        if overdue_months > 0:
            description = f"连续逾期{overdue_months}个月，累计逾期{overdue_days}天"
        else:
            description = "申报及时，无逾期记录"

        return RiskIndicator(
            type="filing_timeliness",
            value=overdue_months,
            benchmark=0,
            deviation=overdue_days,
            score=score,
            level=level,
            description=description
        )

    def calc_preference_validity(
        self,
        qualification_type: str,
        expiry_date: Optional[str],
        is_matching: bool = True
    ) -> RiskIndicator:
        """
        计算优惠备案有效性

        参数：
            qualification_type: 优惠类型（small_scale/high_tech/software/rd/disabled）
            expiry_date: 到期日期（YYYY-MM-DD格式）
            is_matching: 备案与实际是否相符

        返回：
            RiskIndicator对象
        """
        score = 0
        level = "green"
        issues = []

        # 检查资质类型
        valid_types = ["small_scale", "high_tech", "software", "rd", "disabled"]
        if qualification_type not in valid_types:
            return RiskIndicator(
                type="preference_validity",
                value=0,
                benchmark=1,
                deviation=0,
                score=0,
                level="green",
                description="无特殊优惠资质或无需备案"
            )

        # 检查到期日
        if expiry_date:
            try:
                expiry = datetime.strptime(expiry_date, "%Y-%m-%d").date()
                today = date.today()
                days_to_expiry = (expiry - today).days

                if days_to_expiry < 0:
                    score = 50
                    level = "red"
                    issues.append(f"资质已过期{abs(days_to_expiry)}天")
                elif days_to_expiry < 90:
                    score = max(score, 20)
                    level = "orange" if level != "red" else "red"
                    issues.append(f"资质即将在{days_to_expiry}天后到期")
            except ValueError:
                issues.append("资质到期日期格式错误")

        # 检查是否与实际匹配
        if not is_matching:
            score = max(score, 30)
            level = "orange" if level not in ["red"] else "red"
            issues.append("备案与实际情况不符")

        if not issues:
            description = f"{qualification_type}资质有效且备案匹配"
        else:
            description = "；".join(issues)

        return RiskIndicator(
            type="preference_validity",
            value=1 if is_matching else 0,
            benchmark=1,
            deviation=0,
            score=score,
            level=level,
            description=description
        )

    def calc_comprehensive_score(self, indicators: List[RiskIndicator]) -> Dict:
        """
        计算综合风险评分

        参数：
            indicators: 风险指标列表

        返回：
            包含综合评分、风险等级和风险清单的字典
        """
        # 权重配置
        weights = {
            "tax_burden_vat": 0.15,
            "tax_burden_eit": 0.15,
            "invoice_compliance": 0.25,
            "filing_timeliness": 0.20,
            "preference_validity": 0.15,
            "other": 0.10,
        }

        weighted_sum = 0.0
        total_weight = 0.0

        for indicator in indicators:
            weight = weights.get(indicator.type, weights["other"])
            weighted_sum += indicator.score * weight
            total_weight += weight

        # 归一化
        total_score = weighted_sum / total_weight if total_weight > 0 else 0

        # 确定风险等级
        if total_score >= 70:
            risk_level = "红色"
            color = "red"
        elif total_score >= 40:
            risk_level = "橙色"
            color = "orange"
        elif total_score >= 20:
            risk_level = "黄色"
            color = "yellow"
        else:
            risk_level = "绿色"
            color = "green"

        # 生成风险清单
        risk_list = []
        basis_map = {
            "tax_burden_vat": "《增值税暂行条例》第四条",
            "tax_burden_eit": "《企业所得税法》及其实施条例",
            "invoice_compliance": "《发票管理办法》第二十二条、第二十三条",
            "filing_timeliness": "《税收征收管理法》第二十五条",
            "preference_validity": "《税收减免管理办法》",
        }
        suggestion_map = {
            "tax_burden_vat": "核实进项发票取得情况，确认进项抵扣是否充分",
            "tax_burden_eit": "核查成本费用列支是否合规，确认应税收入是否完整",
            "invoice_compliance": "检查发票管理规范，加强发票审核流程",
            "filing_timeliness": "立即完成逾期申报，以后各期确保按时申报",
            "preference_validity": "更新资质备案或确认实际经营情况与备案一致",
        }

        for ind in indicators:
            if ind.score > 0:
                risk_list.append({
                    "risk_id": self._generate_risk_id(),
                    "type": ind.type,
                    "level": ind.level,
                    "description": ind.description,
                    "basis": basis_map.get(ind.type, "相关税法规定"),
                    "suggestion": suggestion_map.get(ind.type, "请核实相关情况"),
                    "score": ind.score
                })

        return {
            "overall_score": round(total_score, 1),
            "risk_level": risk_level,
            "risk_level_color": color,
            "risk_list": risk_list,
            "evaluation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    def generate_report(
        self,
        customer_id: str,
        vat_amount: float,
        sales: float,
        eit_amount: float,
        profit: float,
        industry: str,
        voided_count: int = 0,
        total_invoice_count: int = 0,
        red_invoice_months: int = 0,
        abnormal_amount_count: int = 0,
        overdue_months: int = 0,
        overdue_days: int = 0,
        qualification_type: str = "",
        expiry_date: Optional[str] = None,
        is_matching: bool = True
    ) -> Dict:
        """
        生成完整的税务风险报告

        参数：
            customer_id: 客户编号
            vat_amount: 增值税实缴税额
            sales: 销售收入
            eit_amount: 企业所得税税额
            profit: 利润总额
            industry: 行业
            voided_count: 作废发票份数
            total_invoice_count: 开具发票总份数
            red_invoice_months: 连续红字发票月数
            abnormal_amount_count: 异常金额发票数
            overdue_months: 连续逾期月数
            overdue_days: 累计逾期天数
            qualification_type: 优惠资质类型
            expiry_date: 资质到期日
            is_matching: 备案是否匹配

        返回：
            完整的税务风险报告字典
        """
        indicators = []

        # 1. 增值税税负率
        indicators.append(self.calc_burden_deviation(vat_amount, sales, industry, "vat"))

        # 2. 企业所得税税负率
        if profit > 0:
            indicators.append(self.calc_burden_deviation(eit_amount, profit, industry, "eit"))

        # 3. 发票合规性
        if total_invoice_count > 0:
            indicators.append(self.calc_invoice_compliance(
                voided_count, total_invoice_count,
                red_invoice_months, abnormal_amount_count
            ))

        # 4. 申报及时性
        indicators.append(self.calc_filing_timeliness(overdue_months, overdue_days))

        # 5. 优惠备案有效性
        if qualification_type:
            indicators.append(self.calc_preference_validity(
                qualification_type, expiry_date, is_matching
            ))

        # 计算综合评分
        result = self.calc_comprehensive_score(indicators)
        result["customer_id"] = customer_id
        result["indicators"] = [asdict(ind) for ind in indicators]

        return result


def demo():
    """演示用法"""
    calculator = TaxRiskCalculator()

    # 示例：某商贸企业税务风险评估
    report = calculator.generate_report(
        customer_id="C001",
        vat_amount=50000,
        sales=2000000,
        eit_amount=15000,
        profit=300000,
        industry="百货零售",
        voided_count=15,
        total_invoice_count=200,
        red_invoice_months=2,
        abnormal_amount_count=2,
        overdue_months=0,
        overdue_days=0,
        qualification_type="small_scale",
        expiry_date="2026-12-31",
        is_matching=True
    )

    print("=" * 60)
    print(f"税务风险评估报告 - {report['customer_id']}")
    print("=" * 60)
    print(f"综合评分: {report['overall_score']} 分")
    print(f"风险等级: {report['risk_level']} ({report['risk_level_color']})")
    print(f"评估时间: {report['evaluation_time']}")
    print("-" * 60)
    print("风险清单:")
    for risk in report["risk_list"]:
        print(f"  [{risk['level']}] {risk['description']}")
        print(f"         依据: {risk['basis']}")
        print(f"         建议: {risk['suggestion']}")
    print("=" * 60)

    # 输出JSON格式
    print("\nJSON格式输出:")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    demo()
