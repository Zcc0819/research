#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金山办公取数底稿填写。只写黄格（手填），遇公式格跳过。单位：百万人民币。"""
import datetime
from openpyxl import load_workbook

PATH = "2_2026-08_金山办公_取数底稿.xlsx"
wb = load_workbook(PATH)
TODAY = datetime.date.today().isoformat()
COLS = ["B", "C", "D", "E", "F"]  # FY2021..FY2025
LOG = []


def put(ws, cell, val, log=None):
    c = ws[cell]
    if isinstance(c.value, str) and c.value.startswith("="):
        print("跳过公式格", ws.title, cell)
        return
    c.value = val
    if log:
        LOG.append(log)


def put_row(ws, row, vals, label, src, note=""):
    for col, v, fy in zip(COLS, vals, range(2021, 2026)):
        if v is None:
            continue
        put(ws, f"{col}{row}", v)
    LOG.append((f"{ws.title[0:2]}-{label}-FY2021~25", "见备注", src, note))


# ---------- Sheet1 分部拆解 ----------
s1 = wb["1_分部拆解"]
labels = [
    "个人订阅→WPS个人业务（FY24起新口径）",
    "机构订阅及服务→WPS365业务（FY24起新口径，不可比）",
    "机构授权→WPS软件业务（FY24起新口径）",
    "广告及其他→其他（FY24起新口径）",
]
seg_rev = [
    [1465.30, 2050.13, 2653.35, 3283.29, 3625.54],
    [446.26, 691.99, 957.43, 436.50, 719.90],
    [962.45, 835.62, 655.15, 1267.95, 1461.13],
    [405.60, 306.45, 287.92, 131.72, 120.14],
]
seg_gm = [
    [0.8279, 0.8075, 0.8149, 0.8119, 0.8096],
    [0.9487, 0.9339, 0.9270, 0.9116, 0.9149],
    [0.9604, 0.9343, 0.9312, 0.9272, 0.9502],
    [0.7146, 0.7162, 0.7818, 0.9130, 0.9394],
]
for i, (lab, rev) in enumerate(zip(labels, seg_rev)):
    s1[f"A{i+4}"] = lab
    s1[f"A{i+13}"] = lab  # 占比区行名
    s1[f"A{i+21}"] = lab  # 增速区行名
    s1[f"A{i+29}"] = lab  # 毛利率区行名
    put_row(s1, i + 4, rev, f"分部收入-{lab[:6]}",
            "各年年报「主营业务分产品情况」：2021 p39-40 / 2022 p40-41 / 2023 p38-39 / 2024 p36-37 / 2025 p36",
            "FY2024起公司改分类口径（WPS个人/WPS365/WPS软件/其他），与旧口径按业务实质逐行对应但非严格重述，跨口径增速仅供参考[年报披露]")
    put_row(s1, i + 29, seg_gm[i], f"分部毛利率-{lab[:6]}", "同上分产品表毛利率列", "[年报披露]")
# 清掉多余两行占位（水务科目）
for r in (8, 9, 17, 18, 25, 26, 33, 34):
    s1[f"A{r}"] = "—"

# ---------- Sheet2 利润表五年 ----------
s2 = wb["2_利润表五年"]
rows2 = {
    2: ([3280, 3885, 4556, 5121, 5929], "营业收入", "2025年报p36（592,874万）；余年AKShare利润表与各年年报MD&A一致", "[年报披露]"),
    3: ([2851, 3302, 3886, 4360, 5096], "毛利", "营收-营业成本（营业成本429/583/670/761/833）", "[推算]毛利=营收-营业成本"),
    4: ([1021, 1210, 1411, 1394, 1614], "销售及行政开支", "AKShare利润表 SALE_EXPENSE+MANAGE_EXPENSE", "[推算]销售+管理费用加总"),
    5: ([-17, -13, -95, -16, 0], "财务费用", "AKShare利润表 FINANCE_EXPENSE", "负数=利息净收入[年报披露]"),
    6: ([-15, 73, 34, 79, 166], "应占合联营损益", "AKShare利润表 INVEST_JOINT_INCOME", "[年报披露]"),
    7: ([1119, 1198, 1390, 1750, 1953], "除税前溢利", "AKShare利润表 TOTAL_PROFIT", "[年报披露]"),
    8: ([1069, 1131, 1314, 1655, 1822], "净利润", "AKShare利润表 NETPROFIT", "[年报披露]"),
    9: ([1041, 1118, 1318, 1645, 1836], "归母净利润", "2025年报p36（183,647万）；余年AKShare PARENT_NETPROFIT", "[年报披露]"),
    10: ([201, 179, 56, 89, 33], "非经常项目", "归母-扣非（扣非840/939/1262/1556/1803）", "[推算]=归母-扣非归母"),
}
for r, (vals, lab, src, note) in rows2.items():
    put_row(s2, r, vals, lab, src, note)

# ---------- Sheet3 现金流照妖镜 ----------
s3 = wb["3_现金流照妖镜"]
rows3 = {
    2: ([1864, 1603, 2058, 2196, 2502], "经营活动现金流净额", "AKShare现金流量表 NETCASH_OPERATE", "[年报披露]"),
    3: ([1069, 1131, 1314, 1655, 1822], "净利润", "同Sheet2", "[年报披露]"),
    4: ([-149, -176, -166, -175, -98], "资本开支", "AKShare CONSTRUCT_LONG_ASSET，填负数", "[年报披露]轻资产：资本开支不足OCF的10%"),
    5: ([-1364, 4502, -5210, -4274, -1967], "投资活动现金流净额", "AKShare NETCASH_INVEST", "大额波动系理财/结构性存款配置，非产能投资[年报披露]"),
    6: ([-247, -354, -385, -501, -735], "筹资活动现金流净额", "AKShare NETCASH_FINANCE", "FY2025含收购少数股东股权流出（2025年报p36）[年报披露]"),
    7: ([-277, -323, -337, -406, -482], "已派股息", "AKShare ASSIGN_DIVIDEND_PORFIT，填负数", "分配股利利润偿付利息口径[年报披露]"),
    8: ([1251, 7013, 3474, 895, 691], "期末现金及等价物", "AKShare END_CASH", "现金低≠缺钱：大部分资金在交易性金融资产/定期存款[年报披露]"),
}
for r, (vals, lab, src, note) in rows3.items():
    put_row(s3, r, vals, lab, src, note)

# ---------- Sheet4 资产负债偿债 ----------
s4 = wb["4_资产负债偿债"]
s4["A9"] = "合同负债（预收订阅款，越大越好）★"
rows4 = {
    2: ([10426, 12058, 13974, 15913, 18156], "总资产", "2025年报p106资产总计181.56亿；余年AKShare", "[年报披露]"),
    3: ([7720, 8724, 9949, 11356, 12853], "归母权益", "AKShare TOTAL_PARENT_EQUITY", "[年报披露]"),
    4: ([1254, 7015, 3478, 898, 694], "货币资金", "AKShare MONETARYFUNDS", "另有大额交易性金融资产（理财）不在此行[年报披露]"),
    5: ([0, 0, 0, 0, 0], "短期借贷", "资产负债表无短期借款科目", "[年报披露]零有息负债"),
    6: ([0, 0, 0, 0, 0], "长期借贷", "资产负债表无长期借款科目", "[年报披露]"),
    7: ([0, 0, 0, 0, 0], "永续证券", "无", "[年报披露]"),
    8: ([415, 500, 569, 510, 641], "贸易应收款", "AKShare NOTE_ACCOUNTS_RECE（应收票据及账款）", "[年报披露]"),
    9: ([1421, 1731, 1892, 2136, 2599], "合同负债", "AKShare CONTRACT_LIAB", "订阅制预收款，是收入的先行指标[年报披露]"),
}
for r, (vals, lab, src, note) in rows4.items():
    put_row(s4, r, vals, lab, src, note)

# ---------- Sheet5 同行快照（2026-08-07 东财 snapshot，同一时点） ----------
s5 = wb["5_估值对比"]
peers = [
    ("金山办公", "688111.SH", 265.65, 1232.65, 8.42, 14.04, 15.70, 0.47),
    ("用友网络", "600588.SH", 11.72, 400.47, 5.81, None, -10.28, None),
    ("泛微网络", "603039.SH", 45.58, 140.22, 5.33, 56.42, 2.42, 0.28),
    ("福昕软件", "688095.SH", 69.54, 63.53, 2.44, 140.97, 0.43, 0.43),
    ("科大讯飞", "002230.SZ", 43.53, 1045.63, 4.63, None, -0.91, 0.23),
    ("万兴科技", "300624.SZ", 56.28, 108.81, 11.14, None, -4.80, None),
]
# 表头探测：第1行是列名
hdr = [s5.cell(row=1, column=c).value for c in range(1, 12)]
print("Sheet5表头:", hdr)
for i, p in enumerate(peers):
    r = i + 2
    s5.cell(row=r, column=1).value = p[0]
    s5.cell(row=r, column=2).value = p[1]
    for j, v in enumerate(p[2:], start=3):
        cell = s5.cell(row=r, column=j)
        if isinstance(cell.value, str) and str(cell.value).startswith("="):
            continue
        cell.value = v
LOG.append(("S5-同行快照-全表", "见表", "东财snapshot 2026-08-07收盘 px_snapshot_peers.csv", "PE/ROE为东财口径（TTM、含少数股东）；负PE留空表示亏损；金山东财PE 14.04疑异常（自算静态67），以5b自算为准"))

# ---------- Sheet5b B区 每股指标 ----------
s5b = wb["5b_股价与估值"]
rows5b = {
    20: ([461.0, 461.4, 461.7, 462.2, 463.2], "总股本（百万股）", "由各年分红总额÷每股派息反推（2025年报p63：580,017,310.38÷1.25228）", "[推算]"),
    21: ([16.74, 18.91, 21.55, 24.57, 27.75], "BVPS", "归母权益÷总股本", "[推算]"),
    22: ([2.26, 2.42, 2.85, 3.56, 3.96], "EPS", "归母净利÷总股本", "[推算]与年报披露基本EPS一致"),
    23: ([0.70, 0.70, 0.88, 0.90, 1.2523], "DPS", "各年年报利润分配预案（2025年报p63：每10股12.5228元）", "[年报披露]"),
    24: ([1, 1, 1, 1, 1], "汇率", "A股人民币计价，无汇率问题", "—"),
}
for r, (vals, lab, src, note) in rows5b.items():
    put_row(s5b, r, vals, lab, src, note)

# ---------- Sheet6 取数日志 ----------
s6 = wb["6_取数日志"]
r = 2
while s6.cell(row=r, column=1).value:
    r += 1
for item in LOG:
    s6.cell(row=r, column=1).value = TODAY
    s6.cell(row=r, column=2).value = item[0]
    s6.cell(row=r, column=3).value = item[1]
    s6.cell(row=r, column=4).value = item[2]
    s6.cell(row=r, column=5).value = item[3]
    r += 1

wb.save(PATH)
print("已保存，日志条数:", len(LOG))

# 复核抽查
wb2 = load_workbook(PATH)
print("抽查 S2营收FY2025:", wb2["2_利润表五年"]["F2"].value)
print("抽查 S4合同负债FY2025:", wb2["4_资产负债偿债"]["F9"].value)
print("抽查 S1分部1 FY2021:", wb2["1_分部拆解"]["B4"].value)
print("抽查 S5b DPS FY2025:", wb2["5b_股价与估值"]["F23"].value)
