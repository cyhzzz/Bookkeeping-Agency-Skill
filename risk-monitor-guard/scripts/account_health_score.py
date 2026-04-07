"""
账务健康度评分脚本
综合评估企业账务规范性

输入：
- 凭证列表（从凭证系统）
- 银行对账单
- 往来明细

输出：
- 健康度评分（0-100分）
- 问题分类汇总
- 改进建议

评分维度：
1. 凭证规范性（30分）
   - 要素完整性：日期/摘要/科目/金额/附件
   - 签字完备性

2. 勾稽关系（40分）
   - 资产负债表与利润表勾稽
   - 银行账与现金账勾稽
   - 往来账与客户供应商对账

3. 异常交易检测（30分）
   - 大额异常交易（单笔>50万）
   - 频繁整数交易（>5次/月的整数）
   - 关联交易异常
   - 资金回流检测
"""

import json
import math
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime, date
from collections import defaultdict


@dataclass
class VoucherIssue:
    """凭证问题"""
    voucher_id: str
    issue_type: str
    description: str
    severity: str  # critical/warning/minor


@dataclass
class ReconciliationResult:
    """勾稽结果"""
    item: str
    status: str  # pass/warning/fail
    details: str
    difference: float = 0.0


@dataclass
class AbnormalTransaction:
    """异常交易"""
    transaction_id: str
    transaction_type: str
    amount: float
    date: str
    description: str
    severity: str


class AccountHealthScorer:
    """账务健康度评分器"""

    def __init__(self):
        self.issue_counter = 0
        self.txn_counter = 0

    def _generate_issue_id(self) -> str:
        self.issue_counter += 1
        return f"ISSU-{self.issue_counter:04d}"

    def _generate_txn_id(self) -> str:
        self.txn_counter += 1
        return f"TXN-{self.txn_counter:04d}"

    def check_voucher_completeness(
        self,
        vouchers: List[Dict]
    ) -> Tuple[int, List[VoucherIssue], Dict]:
        """
        检查凭证要素完整性

        参数：
            vouchers: 凭证列表，每项包含：
                - voucher_id: 凭证编号
                - date: 日期
                - summary: 摘要
                - account_code: 科目代码
                - account_name: 科目名称
                - debit_amount: 借方金额
                - credit_amount: 贷方金额
                - attachment_count: 附件数量
                - preparer_sign: 制单人签字
                - reviewer_sign: 审核人签字
                - approver_sign: 批准人签字

        返回：
            (得分, 问题列表, 统计信息)
        """
        total = len(vouchers)
        if total == 0:
            return 0, [], {"total": 0, "complete": 0, "completeness_rate": 0}

        complete_count = 0
        issues = []
        stats = {
            "missing_date": 0,
            "missing_summary": 0,
            "missing_account": 0,
            "missing_amount": 0,
            "missing_attachment": 0,
            "missing_signature": 0,
        }

        for v in vouchers:
            voucher_issues = []

            # 检查日期
            if not v.get("date"):
                stats["missing_date"] += 1
                voucher_issues.append(("missing_date", "缺少日期"))

            # 检查摘要
            if not v.get("summary") or len(v.get("summary", "")) < 3:
                stats["missing_summary"] += 1
                voucher_issues.append(("missing_summary", "摘要不完整"))

            # 检查科目
            if not v.get("account_code") or not v.get("account_name"):
                stats["missing_account"] += 1
                voucher_issues.append(("missing_account", "科目信息缺失"))

            # 检查金额
            debit = v.get("debit_amount", 0)
            credit = v.get("credit_amount", 0)
            if debit == 0 and credit == 0:
                stats["missing_amount"] += 1
                voucher_issues.append(("missing_amount", "金额为零"))

            # 检查附件
            if v.get("attachment_count", 0) == 0:
                stats["missing_attachment"] += 1
                voucher_issues.append(("missing_attachment", "缺少附件"))

            # 检查签字
            signatures = [v.get("preparer_sign"), v.get("reviewer_sign"), v.get("approver_sign")]
            if not all(signatures):
                stats["missing_signature"] += 1
                voucher_issues.append(("missing_signature", "签字不完整"))

            # 记录问题
            if voucher_issues:
                for issue_type, desc in voucher_issues:
                    severity = "critical" if issue_type in ["missing_amount", "missing_account"] else "warning"
                    issues.append(VoucherIssue(
                        voucher_id=v.get("voucher_id", "UNKNOWN"),
                        issue_type=issue_type,
                        description=desc,
                        severity=severity
                    ))
            else:
                complete_count += 1

        completeness_rate = complete_count / total
        # 得分：完整性 * 30分
        score = completeness_rate * 30

        detail_stats = {
            "total": total,
            "complete": complete_count,
            "completeness_rate": completeness_rate,
            "missing_date": stats["missing_date"],
            "missing_summary": stats["missing_summary"],
            "missing_account": stats["missing_account"],
            "missing_attachment": stats["missing_attachment"],
            "missing_signature": stats["missing_signature"],
        }

        return round(score, 1), issues, detail_stats

    def check_voucher_sequence(self, voucher_ids: List[str]) -> List[VoucherIssue]:
        """
        检查凭证编号连续性

        参数：
            voucher_ids: 凭证编号列表

        返回：
            问题列表
        """
        issues = []

        # 提取编号中的数字部分
        def extract_num(vid):
            import re
            nums = re.findall(r'\d+', str(vid))
            return int(nums[-1]) if nums else 0

        if not voucher_ids:
            return issues

        sorted_ids = sorted(voucher_ids, key=extract_num)
        expected = extract_num(sorted_ids[0])

        for vid in sorted_ids[1:]:
            num = extract_num(vid)
            if num != expected + 1:
                issues.append(VoucherIssue(
                    voucher_id=vid,
                    issue_type="sequence_break",
                    description=f"凭证编号不连续，期望{expected+1}，实际{num}",
                    severity="warning"
                ))
            expected = num

        return issues

    def check_reconciliation(
        self,
        balance_sheet: Dict,
        income_statement: Dict,
        bank_statements: List[Dict],
        cash_balance: float,
        receivables: Dict,
        payables: Dict
    ) -> Tuple[int, List[ReconciliationResult], Dict]:
        """
        勾稽关系校验

        参数：
            balance_sheet: 资产负债表数据
                - undistributed_profit_begin: 期初未分配利润
                - undistributed_profit_end: 期末未分配利润
                - net_profit: 净利润
                - bank_cash: 银行存款
                - cash_on_hand: 现金
            income_statement: 利润表数据
                - net_profit: 净利润
            bank_statements: 银行对账单列表
                - account_name: 账户名称
                - ending_balance: 期末余额
            cash_balance: 现金实盘金额
            receivables: 应收账款（客户对账）
                - book_balance: 账面余额
                - confirmed_balance: 已对账确认余额
            payables: 应付账款（供应商对账）
                - book_balance: 账面余额
                - confirmed_balance: 已对账确认余额

        返回：
            (得分, 勾稽结果列表, 统计信息)
        """
        results = []
        total_checks = 4
        passed_checks = 0

        # 1. 资产负债表与利润表勾稽
        # 期末未分配利润 - 期初未分配利润 = 本期净利润（允许±1元误差）
        bs_profit_diff = (
            balance_sheet.get("undistributed_profit_end", 0) -
            balance_sheet.get("undistributed_profit_end", 0)
        )
        is_net_profit = income_statement.get("net_profit", 0)
        diff = abs(bs_profit_diff - is_net_profit)

        if diff <= 1:
            results.append(ReconciliationResult(
                item="资产负债表-利润表勾稽",
                status="pass",
                details=f"未分配利润变动({bs_profit_diff})与净利润({is_net_profit})一致",
                difference=diff
            ))
            passed_checks += 1
        else:
            results.append(ReconciliationResult(
                item="资产负债表-利润表勾稽",
                status="fail",
                details=f"未分配利润变动({bs_profit_diff})与净利润({is_net_profit})差异{diff}元",
                difference=diff
            ))

        # 2. 银行账与银行对账单勾稽
        bank_cash = balance_sheet.get("bank_cash", 0)
        total_bank_statement = sum(s.get("ending_balance", 0) for s in bank_statements)
        bank_diff = abs(bank_cash - total_bank_statement)

        if bank_diff == 0:
            results.append(ReconciliationResult(
                item="银行账与银行对账单勾稽",
                status="pass",
                details=f"银行账余额({bank_cash})与对账单合计({total_bank_statement})一致",
                difference=bank_diff
            ))
            passed_checks += 1
        else:
            results.append(ReconciliationResult(
                item="银行账与银行对账单勾稽",
                status="fail",
                details=f"银行账余额({bank_cash})与对账单合计({total_bank_statement})存在差异{bank_diff}元",
                difference=bank_diff
            ))

        # 3. 现金账与实盘勾稽
        cash_on_hand = balance_sheet.get("cash_on_hand", 0)
        cash_diff = abs(cash_on_hand - cash_balance)

        if cash_diff == 0:
            results.append(ReconciliationResult(
                item="现金账与实盘勾稽",
                status="pass",
                details=f"现金账余额({cash_on_hand})与实盘金额({cash_balance})一致",
                difference=cash_diff
            ))
            passed_checks += 1
        else:
            results.append(ReconciliationResult(
                item="现金账与实盘勾稽",
                status="fail",
                details=f"现金账余额({cash_on_hand})与实盘金额({cash_balance})存在差异{cash_diff}元",
                difference=cash_diff
            ))

        # 4. 往来账核对
        ar_diff = abs(receivables.get("book_balance", 0) - receivables.get("confirmed_balance", 0))
        ap_diff = abs(payables.get("book_balance", 0) - payables.get("confirmed_balance", 0))

        if ar_diff == 0 and ap_diff == 0:
            results.append(ReconciliationResult(
                item="往来账与客户供应商对账",
                status="pass",
                details="应收账款、应付账款均已与客户供应商完成对账确认",
                difference=0
            ))
            passed_checks += 1
        else:
            details_parts = []
            if ar_diff > 0:
                details_parts.append(f"应收账款差异{ar_diff}元")
            if ap_diff > 0:
                details_parts.append(f"应付账款差异{ap_diff}元")
            results.append(ReconciliationResult(
                item="往来账与客户供应商对账",
                status="fail",
                details="；".join(details_parts),
                difference=ar_diff + ap_diff
            ))

        # 计算得分
        score = (passed_checks / total_checks) * 40

        stats = {
            "total_checks": total_checks,
            "passed_checks": passed_checks,
            "pass_rate": passed_checks / total_checks,
            "items": [
                {"item": r.item, "status": r.status, "difference": r.difference}
                for r in results
            ]
        }

        return round(score, 1), results, stats

    def detect_abnormal_transactions(
        self,
        transactions: List[Dict],
        monthly_avg_sales: float = 0
    ) -> Tuple[int, List[AbnormalTransaction], Dict]:
        """
        异常交易检测

        参数：
            transactions: 交易列表，每项包含：
                - transaction_id: 交易编号
                - date: 交易日期
                - amount: 交易金额
                - description: 交易描述
                - counterparty: 交易对手
                - is_related_party: 是否关联方
            monthly_avg_sales: 月均销售额（用于判断大额交易）

        返回：
            (得分, 异常交易列表, 统计信息)
        """
        if not transactions:
            return 30, [], {"total": 0, "abnormal_count": 0, "abnormal_rate": 0}

        abnormal_txns = []
        large_amount_threshold = monthly_avg_sales * 0.5 if monthly_avg_sales > 0 else 500000

        stats = {
            "large_amount_count": 0,
            "frequent_integer_count": 0,
            "related_party_count": 0,
            "fund_reflow_count": 0,
        }

        # 按日期排序
        sorted_txns = sorted(transactions, key=lambda x: x.get("date", ""))

        # 检测大额异常交易
        for txn in transactions:
            amount = txn.get("amount", 0)
            if amount > large_amount_threshold:
                abnormal_txns.append(AbnormalTransaction(
                    transaction_id=txn.get("transaction_id", ""),
                    transaction_type="large_amount",
                    amount=amount,
                    date=txn.get("date", ""),
                    description=f"单笔交易金额{amount}元超过月均销售额50%({large_amount_threshold}元)",
                    severity="warning"
                ))
                stats["large_amount_count"] += 1

        # 检测频繁整数交易（连续5笔以上整数）
        integer_txns = []
        for txn in transactions:
            amount = txn.get("amount", 0)
            if amount == int(amount) and amount >= 10000:  # 万元以上整数
                integer_txns.append(txn)

        if len(integer_txns) >= 5:
            # 找出连续整数交易
            consecutive = 1
            for i in range(1, len(integer_txns)):
                if integer_txns[i].get("amount") == integer_txns[i-1].get("amount"):
                    consecutive += 1
                    if consecutive >= 5:
                        for j in range(i-4, i+1):
                            if integer_txns[j].get("transaction_id") not in [t.transaction_id for t in abnormal_txns]:
                                abnormal_txns.append(AbnormalTransaction(
                                    transaction_id=integer_txns[j].get("transaction_id", ""),
                                    transaction_type="frequent_integer",
                                    amount=integer_txns[j].get("amount", 0),
                                    date=integer_txns[j].get("date", ""),
                                    description=f"连续5笔以上整数交易，金额{integer_txns[j].get('amount')}元",
                                    severity="warning"
                                ))
                                stats["frequent_integer_count"] += 1
                else:
                    consecutive = 1

        # 检测关联交易异常
        for txn in transactions:
            if txn.get("is_related_party", False):
                amount = txn.get("amount", 0)
                # 简化：关联交易超过50万标记
                if amount > 500000:
                    abnormal_txns.append(AbnormalTransaction(
                        transaction_id=txn.get("transaction_id", ""),
                        transaction_type="related_party",
                        amount=amount,
                        date=txn.get("date", ""),
                        description=f"关联方大额交易{amount}元，需关注交易价格公允性",
                        severity="critical"
                    ))
                    stats["related_party_count"] += 1

        # 简化：资金回流检测（收款后24小时内原路返回）
        # 按金额、时间、对手方匹配
        receipt_map = {}  # (amount, counterparty) -> [(date, txn_id)]
        for txn in transactions:
            if txn.get("amount", 0) > 0:
                key = (txn.get("amount"), txn.get("counterparty"))
                if key not in receipt_map:
                    receipt_map[key] = []
                receipt_map[key].append((txn.get("date"), txn.get("transaction_id")))

        for txn in transactions:
            if txn.get("amount", 0) < 0:  # 支出
                key = (-txn.get("amount"), txn.get("counterparty"))
                if key in receipt_map:
                    for receipt_date, receipt_id in receipt_map[key]:
                        # 简化：同日期即视为回流风险
                        if receipt_date == txn.get("date"):
                            abnormal_txns.append(AbnormalTransaction(
                                transaction_id=txn.get("transaction_id", ""),
                                transaction_type="fund_reflow",
                                amount=txn.get("amount", 0),
                                date=txn.get("date", ""),
                                description=f"疑似资金回流，与收据{receipt_id}金额相同、对手相同",
                                severity="critical"
                            ))
                            stats["fund_reflow_count"] += 1

        # 计算异常比例
        abnormal_count = len(abnormal_txns)
        abnormal_rate = abnormal_count / len(transactions) if transactions else 0

        # 得分：30分满分，异常率越高扣分越多
        if abnormal_rate == 0:
            score = 30
        elif abnormal_rate < 0.05:
            score = 25
        elif abnormal_rate < 0.10:
            score = 20
        elif abnormal_rate < 0.20:
            score = 15
        else:
            score = 10

        stats["total"] = len(transactions)
        stats["abnormal_count"] = abnormal_count
        stats["abnormal_rate"] = abnormal_rate

        return score, abnormal_txns, stats

    def calc_health_score(
        self,
        voucher_score: float,
        reconciliation_score: float,
        abnormal_score: float
    ) -> Tuple[int, str, Dict]:
        """
        计算综合健康度评分

        参数：
            voucher_score: 凭证规范性得分（0-30）
            reconciliation_score: 勾稽关系得分（0-40）
            abnormal_score: 异常交易得分（0-30）

        返回：
            (总分, 等级, 明细)
        """
        total = voucher_score + reconciliation_score + abnormal_score

        if total >= 90:
            level = "优秀"
        elif total >= 75:
            level = "良好"
        elif total >= 60:
            level = "合格"
        else:
            level = "不合格"

        breakdown = {
            "voucher_score": voucher_score,
            "voucher_max": 30,
            "reconciliation_score": reconciliation_score,
            "reconciliation_max": 40,
            "abnormal_score": abnormal_score,
            "abnormal_max": 30,
            "total_score": total,
            "level": level
        }

        return round(total), level, breakdown

    def generate_report(
        self,
        customer_id: str,
        vouchers: List[Dict],
        balance_sheet: Dict,
        income_statement: Dict,
        bank_statements: List[Dict],
        cash_balance: float,
        receivables: Dict,
        payables: Dict,
        transactions: List[Dict],
        monthly_avg_sales: float = 0
    ) -> Dict:
        """
        生成完整的账务健康度报告

        参数：
            customer_id: 客户编号
            vouchers: 凭证列表
            balance_sheet: 资产负债表
            income_statement: 利润表
            bank_statements: 银行对账单列表
            cash_balance: 现金实盘金额
            receivables: 应收账款对账信息
            payables: 应付账款对账信息
            transactions: 交易列表
            monthly_avg_sales: 月均销售额

        返回：
            完整的健康度报告
        """
        # 1. 凭证规范性检查
        voucher_score, voucher_issues, voucher_stats = self.check_voucher_completeness(vouchers)
        sequence_issues = self.check_voucher_sequence([v.get("voucher_id") for v in vouchers])
        all_voucher_issues = voucher_issues + sequence_issues

        # 2. 勾稽关系校验
        reconciliation_score, reconciliation_results, reconciliation_stats = self.check_reconciliation(
            balance_sheet, income_statement, bank_statements,
            cash_balance, receivables, payables
        )

        # 3. 异常交易检测
        abnormal_score, abnormal_txns, abnormal_stats = self.detect_abnormal_transactions(
            transactions, monthly_avg_sales
        )

        # 4. 计算综合评分
        total_score, level, breakdown = self.calc_health_score(
            voucher_score, reconciliation_score, abnormal_score
        )

        # 分类汇总问题
        critical_issues = [v for v in all_voucher_issues if v.severity == "critical"]
        warning_issues = [v for v in all_voucher_issues if v.severity == "warning"]

        return {
            "customer_id": customer_id,
            "evaluation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_score": total_score,
            "level": level,
            "breakdown": breakdown,
            "voucher_check": {
                "score": voucher_score,
                "total_count": voucher_stats["total"],
                "complete_count": voucher_stats["complete"],
                "completeness_rate": round(voucher_stats["completeness_rate"] * 100, 1),
                "issues": [asdict(v) for v in all_voucher_issues],
                "critical_count": len(critical_issues),
                "warning_count": len(warning_issues),
            },
            "reconciliation": {
                "score": reconciliation_score,
                "results": [asdict(r) for r in reconciliation_results],
                "passed_checks": reconciliation_stats["passed_checks"],
                "total_checks": reconciliation_stats["total_checks"],
            },
            "abnormal_transactions": {
                "score": abnormal_score,
                "transactions": [asdict(t) for t in abnormal_txns],
                "abnormal_count": abnormal_stats["abnormal_count"],
                "abnormal_rate": round(abnormal_stats["abnormal_rate"] * 100, 1),
            },
            "issues_summary": {
                "critical": [asdict(v) for v in critical_issues],
                "warning": [asdict(v) for v in warning_issues],
                "reconciliation_failures": [
                    asdict(r) for r in reconciliation_results if r.status == "fail"
                ],
                "abnormal_txns": [asdict(t) for t in abnormal_txns],
            },
            "improvement_suggestions": self._generate_suggestions(
                all_voucher_issues, reconciliation_results, abnormal_txns, level
            )
        }

    def _generate_suggestions(
        self,
        voucher_issues: List[VoucherIssue],
        reconciliation_results: List[ReconciliationResult],
        abnormal_txns: List[AbnormalTransaction],
        level: str
    ) -> List[str]:
        """生成改进建议"""
        suggestions = []

        # 基于凭证问题
        if any(v.issue_type == "missing_attachment" for v in voucher_issues):
            suggestions.append("完善凭证附件管理，确保所有凭证都附有完整的原始单据")
        if any(v.issue_type == "missing_signature" for v in voucher_issues):
            suggestions.append("加强凭证审核流程，确保所有凭证经过制单、审核、批准签字")
        if any(v.issue_type == "sequence_break" for v in voucher_issues):
            suggestions.append("检查凭证编号管理，规范编号编制规则，避免断号重号")

        # 基于勾稽问题
        for r in reconciliation_results:
            if r.status == "fail":
                if "未分配利润" in r.item:
                    suggestions.append("核查资产负债表与利润表之间的勾稽关系，确保净利润计算准确")
                elif "银行" in r.item:
                    suggestions.append("核对银行存款日记账与银行对账单，及时处理未达账项")
                elif "现金" in r.item:
                    suggestions.append("进行现金盘点，确保账实相符")
                elif "往来" in r.item:
                    suggestions.append("加强应收账款应付账款管理，定期与客户供应商对账确认")

        # 基于异常交易
        txn_types = set(t.transaction_type for t in abnormal_txns)
        if "large_amount" in txn_types:
            suggestions.append("对大额交易进行专项审查，完善大额交易审批流程")
        if "frequent_integer" in txn_types:
            suggestions.append("避免频繁进行整数金额交易，确保交易真实性")
        if "related_party" in txn_types:
            suggestions.append("规范关联交易定价，确保关联交易价格公允")
        if "fund_reflow" in txn_types:
            suggestions.append("严禁资金回流，确保资金流向真实合规")

        # 整体建议
        if level in ["不合格", "合格"]:
            suggestions.append("建议进行专项账务整顿，全面提升账务管理水平")
        elif level == "良好":
            suggestions.append("继续保持现有管理水平，针对性改进小问题")

        return suggestions


def demo():
    """演示用法"""
    scorer = AccountHealthScorer()

    # 示例数据
    vouchers = [
        {
            "voucher_id": "记-001",
            "date": "2024-01-15",
            "summary": "收到货款",
            "account_code": "1002",
            "account_name": "银行存款",
            "debit_amount": 58500,
            "credit_amount": 0,
            "attachment_count": 2,
            "preparer_sign": "张三",
            "reviewer_sign": "李四",
            "approver_sign": "王五"
        },
        {
            "voucher_id": "记-002",
            "date": "2024-01-20",
            "summary": "支付",
            "account_code": "6602",
            "account_name": "销售费用",
            "debit_amount": 0,
            "credit_amount": 10000,
            "attachment_count": 0,  # 缺少附件
            "preparer_sign": "张三",
            "reviewer_sign": "",   # 缺少审核签字
            "approver_sign": "王五"
        },
    ]

    balance_sheet = {
        "undistributed_profit_begin": 500000,
        "undistributed_profit_end": 620000,
        "bank_cash": 800000,
        "cash_on_hand": 5000,
    }

    income_statement = {
        "net_profit": 120000,
    }

    bank_statements = [
        {"account_name": "基本户", "ending_balance": 800000},
    ]

    transactions = [
        {"transaction_id": "T001", "date": "2024-01-15", "amount": 58500, "counterparty": "A公司", "is_related_party": False},
        {"transaction_id": "T002", "date": "2024-01-20", "amount": -10000, "counterparty": "B公司", "is_related_party": False},
        {"transaction_id": "T003", "date": "2024-01-25", "amount": 600000, "counterparty": "关联方", "is_related_party": True},
    ]

    report = scorer.generate_report(
        customer_id="C001",
        vouchers=vouchers,
        balance_sheet=balance_sheet,
        income_statement=income_statement,
        bank_statements=bank_statements,
        cash_balance=5000,
        receivables={"book_balance": 200000, "confirmed_balance": 200000},
        payables={"book_balance": 150000, "confirmed_balance": 150000},
        transactions=transactions,
        monthly_avg_sales=1000000
    )

    print("=" * 60)
    print(f"账务健康度诊断报告 - {report['customer_id']}")
    print("=" * 60)
    print(f"综合评分: {report['total_score']} 分 ({report['level']})")
    print(f"评估时间: {report['evaluation_time']}")
    print("-" * 60)
    print(f"分项得分:")
    print(f"  凭证规范性: {report['breakdown']['voucher_score']}/{report['breakdown']['voucher_max']}分")
    print(f"  勾稽关系: {report['breakdown']['reconciliation_score']}/{report['breakdown']['reconciliation_max']}分")
    print(f"  异常交易: {report['breakdown']['abnormal_score']}/{report['breakdown']['abnormal_max']}分")
    print("-" * 60)
    print(f"凭证合规性: {report['voucher_check']['completeness_rate']}%")
    print(f"勾稽校验: {report['reconciliation']['passed_checks']}/{report['reconciliation']['total_checks']}项通过")
    print(f"异常交易: {report['abnormal_transactions']['abnormal_count']}笔")
    print("-" * 60)
    print("改进建议:")
    for i, s in enumerate(report['improvement_suggestions'], 1):
        print(f"  {i}. {s}")
    print("=" * 60)

    print("\nJSON格式输出:")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    demo()
