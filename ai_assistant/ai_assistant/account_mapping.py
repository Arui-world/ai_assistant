from __future__ import annotations

import re


ACCOUNT_CATALOG = {
    "1001": "库存现金",
    "1002": "银行存款",
    "1122": "应收账款",
    "1221": "其他应收款",
    "1405": "库存商品",
    "1601": "固定资产",
    "1602": "累计折旧",
    "2202": "应付账款",
    "2211": "应付职工薪酬",
    "2221": "应交税费",
    "2241": "其他应付款",
    "3001": "实收资本",
    "3104": "利润分配",
    "4104": "本年利润",
    "6001": "主营业务收入",
    "6051": "其他业务收入",
    "6401": "主营业务成本",
    "6601": "销售费用",
    "6602": "管理费用",
    "6603": "财务费用",
    "6711": "营业外支出",
    "6801": "所得税",
}

EXPENSE_RULES = [
    (("代发工资手续费", "跨行汇款手续费", "余额变动提醒手续费", "手续费", "收费明细", "对公收费"), "财务费用-手续费"),
    (("服务费", "咨询费", "技术服务"), "管理费用-办公费"),
    (("运费", "物流", "运输"), "管理费用-运输费"),
    (("快递", "办公", "耗材", "用品"), "管理费用-办公费"),
    (("房租", "租金", "物业"), "管理费用-房租"),
    (("工资", "薪资", "薪酬", "代发"), "应付职工薪酬-工资"),
    (("社保", "社会保险"), "应付职工薪酬-社保"),
    (("公积金",), "应付职工薪酬-公积金"),
    (("税", "国库", "税收", "增值税", "所得税"), "应交税费"),
    (("报销", "差旅", "餐费", "招待"), "管理费用-报销"),
    (("货款", "采购", "供应商", "材料"), "应付账款"),
    (("利息",), "财务费用-利息"),
]

INCOME_RULES = [
    (("销售", "收入", "货款", "客户", "回款", "收款"), "主营业务收入"),
    (("利息",), "财务费用-利息收入"),
    (("投资", "股东", "实收资本"), "实收资本"),
    (("退款", "退回", "返还"), "其他应收款"),
]


ACCOUNT_ROOTS = {
    "库存现金": "asset",
    "银行存款": "asset",
    "应收账款": "asset",
    "其他应收款": "asset",
    "库存商品": "asset",
    "固定资产": "asset",
    "累计折旧": "asset_credit",
    "应付账款": "liability",
    "应付职工薪酬": "liability",
    "应交税费": "liability",
    "其他应付款": "liability",
    "实收资本": "equity",
    "利润分配": "equity",
    "本年利润": "equity",
    "主营业务收入": "income",
    "其他业务收入": "income",
    "主营业务成本": "expense",
    "销售费用": "expense",
    "管理费用": "expense",
    "财务费用": "expense",
    "营业外支出": "expense",
    "所得税": "expense",
}


def normalize_text(*parts):
    return " ".join(str(part or "") for part in parts).strip()


def counterparty_suffix(counterparty):
    cleaned = re.sub(r"\s+", "", str(counterparty or ""))
    if not cleaned:
        return "待确认"
    return cleaned[:24]


def base_account(account):
    return str(account or "").split("-", 1)[0]


def account_root_type(account):
    base = base_account(account)
    return ACCOUNT_ROOTS.get(base, "unknown")


def is_valid_account(account):
    if not account:
        return False
    return base_account(account) in ACCOUNT_ROOTS


def allowed_base_accounts():
    return sorted(ACCOUNT_ROOTS)


def map_transaction(summary=None, purpose=None, counterparty=None, direction="out"):
    text = normalize_text(summary, purpose, counterparty)
    compact = re.sub(r"\s+", "", text)

    if direction == "out":
        if compact.isdigit() and len(compact) >= 6:
            return "管理费用-公积金", "银行存款", "数字摘要按公积金/缴费编号归类"
        for keywords, debit_account in EXPENSE_RULES:
            if any(keyword in compact for keyword in keywords):
                return debit_account, "银行存款", f"关键词匹配：{debit_account}"
        if not str(counterparty or "").strip() and any(keyword in compact for keyword in ("手续费", "收费", "扣费", "服务费")):
            return "财务费用-手续费", "银行存款", "对方单位为空且摘要为收费类，按银行手续费归类"
        return f"其他应付款-{counterparty_suffix(counterparty)}", "银行存款", "未命中规则，按往来款处理，需人工复核"

    for keywords, credit_account in INCOME_RULES:
        if any(keyword in compact for keyword in keywords):
            return "银行存款", credit_account, f"关键词匹配：{credit_account}"
    if "网转" in compact or "转账" in compact or "往来" in compact:
        return "银行存款", f"其他应付款-{counterparty_suffix(counterparty)}", "转入往来款规则"
    return "银行存款", f"其他应付款-{counterparty_suffix(counterparty)}", "未命中规则，按收到往来款处理，需人工复核"
