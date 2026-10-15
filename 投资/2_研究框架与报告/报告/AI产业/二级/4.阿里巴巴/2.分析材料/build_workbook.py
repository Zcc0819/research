"""生成阿里巴巴取数底稿（八 Sheet，口径同《取数底稿模板_二级市场》）。

所有数字来源：20-F FY2022–FY2026（EDGAR 一手）＋ 6-K EX-99.1 ＋ AKShare 东财港股接口（二手，已交叉验证）。
比率一律由脚本计算，不手抄，避免算术错误。
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent / "2_2026-09_阿里巴巴_取数底稿.xlsx"

FY = ["FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]

# ───────────────────────── 一手数据（RMB 百万） ─────────────────────────
rev = [853062, 868687, 941168, 996347, 1023670]
gp = [313612, 318992, 354845, 398062, 407534]
selling = [119799, 103496, 115141, 144021, 245023]
admin = [31922, 42183, 41985, 44239, 33082]
rd = [55465, 56744, 52256, 57151, 66533]
imp = [36788, 16218, 32113, 12507, 14594]
op = [69638, 100351, 113350, 140905, 50150]
int_inv = [-15702, -11071, -9964, 20759, 87512]  # 正=净收益
tax = [26815, 15549, 22529, 35445, 30045]
ni = [47079, 65573, 71332, 125976, 102127]
attrib = [62249, 72783, 80009, 130109, 103592]

cloud_rev = [102016, 103497, 106374, 118028, 158132]
cloud_ebita = [3744, 4101, 6121, 10556, 14265]
aidc_rev = [62185, 70506, 102598, 132300, 144170]
aidc_ebita = [-8614, -4944, -8035, -15137, -2051]
cons_ebita = [130397, 147911, 165028, 173065, 76416]

ocf = [142759, 199752, 182593, 163509, 76213]
capex = [53309, 34330, 32087, 85972, 126063]
fcf = [98874, 171663, 156210, 73870, -46609]
da_cf = [48065, 46938, 44504, 42459, 47118]
dep_ppe = [27808, 27799, 26640, 29260, 37067]
ebitda = [158205, 175710, 191668, 202325, 113483]
buyback = [61225, 74746, 88745, 86662, 7638]
dividend = [881, 489, 18492, 29741, 35049]
new_debt = [9427, 22790, 20570, 52788, 84228]
repay = [7128, 11448, 13092, 43678, 78458]
bond_iss = [0, 0, 0, 71644, 33262]
icf = [-198592, -135506, -21824, -185415, -67336]
fcf_fin = [-64449, -65619, -108244, -76215, -20573]

ta = [1695553, 1753044, 1764829, 1804227, 1909570]
tl = [613360, 630123, 652230, 714121, 783300]
eq_p = [958134, 999515, 997272, 1021571, 1068731]
eq_m = [124059, 123406, 115327, 68535, 57539]
ppe = [171806, 176031, 185161, 203348, 282699]
cash = [189898, 193086, 248125, 145487, 131530]
sti = [265187, 331384, 322904, 282606, 185364]
goodwill = [269581, 268091, 259679, 255501, 247378]
st_loan = [8841, 7466, 12749, 22562, 28224]
lt_loan = [38244, 52023, 55686, 49909, 47450]
conv = [0, 0, 0, 35834, 66837]
nwc = [254751, 312615, 331357, 238703, 134371]
ca = [638535, 697966, 752864, 674049, 610769]
cl = [383784, 385351, 421507, 435346, 476398]

# 新口径四分部（FY2026 20-F，仅三年可得）
seg3 = ["FY2024", "FY2025", "FY2026"]
cn_ec_rev = [490101, 508380, 554217]
cn_ec_ebita = [186970, 193223, 107509]
allo_rev = [317539, 338347, 254367]
allo_ebita = [-11252, -9499, -35737]
unalloc_ebita = [-6190, -4337, -5150]
elim_ebita = [-2586, -1741, -2420]
cmr = [307950, 326769, 343867]
quick = [50852, 53588, 78520]

# ───────────────────────── 样式 ─────────────────────────
T = Font(bold=True, size=13, color="FFFFFF")
H = Font(bold=True, size=10)
B = Font(bold=True)
RED = Font(bold=True, color="C00000")
GRN = Font(bold=True, color="006100")
F_T = PatternFill("solid", fgColor="1F4E78")
F_H = PatternFill("solid", fgColor="DDEBF7")
F_S = PatternFill("solid", fgColor="FFF2CC")
F_K = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()


def newsheet(name, title, widths):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = T
    for c in range(1, len(widths) + 1):
        ws.cell(row=1, column=c).fill = F_T
        ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
    ws.freeze_panes = "B2"
    return ws


def put(ws, r, vals, font=None, fill=None, fmt="#,##0"):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BD
        if font:
            c.font = font
        if fill:
            c.fill = fill
        if isinstance(v, (int, float)) and i > 1:
            c.number_format = fmt
    return r + 1


def pct(a, b):
    return None if not b else round(a / b * 100, 1)


def yoy(s):
    return [None] + [pct(s[i] - s[i - 1], abs(s[i - 1])) for i in range(1, len(s))]


# ═════════════════════ Sheet 0_说明 ═════════════════════
ws = newsheet("0_说明", "阿里巴巴集团（09988.HK / BABA）取数底稿 · 说明与口径声明", [26, 100])
rows = [
    ("公司", "阿里巴巴集团控股有限公司 Alibaba Group Holding Limited（09988.HK 港股／BABA 纽交所 ADS／89988 人民币柜台）"),
    ("注册地/架构", "开曼群岛；2019-11-26 港股二次上市，2024-08-28 转为香港双重主要上市；1 ADS = 8 股普通股"),
    ("★ 财年口径", "财年止于 3 月 31 日。FY2026 = 2025-04-01 至 2026-03-31。跨期比较必须先对齐财年，禁止与自然年公司直接比"),
    ("报表币种", "人民币（RMB）百万元。⚠️ 东财接口 CURRENCY 字段错标 HKD（IS_CNY_CODE=0），实际列报为人民币，与 20-F 一致"),
    ("最新年报", "20-F FY2026，2026-05-20 提交（report date 2026-03-31）"),
    ("最新季报", "6-K EX-99.1，June Quarter 2026 = FY2027 Q1，2026-08-20 发布"),
    ("会计准则/审计", "U.S. GAAP；审计师 PricewaterhouseCoopers（香港）。五年 CAM 恒为「商誉减值」，无变化"),
    ("", ""),
    ("★ 口径地雷 1", "分部结构 FY2027Q1 起重组：云智能＋T-Head → AI Cloud and Compute Services；AI Labs/Qwen 从 All Others 拆出。新旧不可比，跨期只能用同一份公告内的两列"),
    ("★ 口径地雷 2", "分部结构在 FY2025 也变过：FY2024 20-F 为七分部（淘天/云/AIDC/菜鸟/本地生活/大文娱/All others），FY2026 20-F 为四分部（中国电商/AIDC/云/All others）。仅「云智能」与「AIDC」两条可做五年连续序列"),
    ("★ 口径地雷 3", "capex 两个口径：现金流量表/Item 5.B「购建固定资产」FY2026=126,063（含办公园区）；FCF 调节表「购建房产设备」FY2026=122,021（剔除办公园区）。差 4,042，引用必须声明"),
    ("★ 口径地雷 4", "AI 收入占比两个分母：CEO 信「AI 占 30%」分母为云外部收入；季报 AI 收入 12,376 分母为分部总收入 48,437（=25.6%）。两者不可连续比较"),
    ("★ 口径地雷 5", "净利润含大额投资收益（FY2026 利息与投资净收益 +87,512），归母净利 103,592 > 经营利润 50,150。给阿里估值禁用净利润/PE 单锚，必须用经营利润或分部 EBITA"),
    ("", ""),
    ("证据分级图例", "[披露]=年报/公告原文｜[推算]=本底稿据披露数计算（比率列全部为推算）｜[估算]=非披露口径的假设（如汇率）｜[管理层口径]=致股东信/业绩会表述｜[第三方]=Gartner/IDC/Omdia/F&S"),
    ("数据源", "① SEC EDGAR：20-F FY2022–FY2026 五份主文档＋Exhibit（一手）② EDGAR 6-K EX-99.1 五份（业绩/配股/CMC List）③ AKShare 东财港股三大报表＋财务指标（二手，已对锚）④ fetch_stock_price.py（东财快照＋腾讯降级日线）"),
    ("交叉验证结论", "锚点三科目双源一致：营业额 FY2026 1,023,670（20-F=东财）、归母净利 103,592、总资产 1,909,570，误差 0%。→ 东财其余科目可信度提升，可用于填 Sheet2/3/4"),
    ("巨潮通道失败记录", "巨潮资讯网检索港股「年報/年度報告/年报」均下载 0 份 → 全量转 EDGAR 通道（英文可检索、分部披露更细、五年齐全）"),
    ("底稿完成日", "2026-09-14｜取价日 2026-09-14｜编制：Noah（个人参数取《投资观点配置_Noah.md》）"),
]
r = 3
for k, v in rows:
    ws.cell(row=r, column=1, value=k).font = B
    c = ws.cell(row=r, column=2, value=v)
    c.alignment = WRAP
    if k.startswith("★"):
        ws.cell(row=r, column=1).fill = F_K
        c.fill = F_K
    r += 1
ws.row_dimensions[1].height = 22

# ═════════════════════ Sheet 1_分部拆解 ═════════════════════
ws = newsheet("1_分部拆解", "分部拆解（→ 打灯 #1 收入质量 / 分析文件第一节）", [40, 14, 14, 14, 14, 14, 52])
r = 3
r = put(ws, r, ["A. 五年连续序列（仅云智能与 AIDC 跨口径可连续）", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["RMB 百万"] + FY + ["出处 / 备注"], H, F_H)
r = put(ws, r, ["★ 云智能集团 收入"] + cloud_rev + ["20-F FY2024 表 t068 + FY2026 表 t068 [披露]"], B)
r = put(ws, r, ["  云收入 YoY %"] + yoy(cloud_rev) + ["[推算] ★ AI 之前三年近乎零增长（+1.5%/+2.8%），FY2025 起才提速"], None, None, "0.0")
r = put(ws, r, ["★ 云智能集团 Adjusted EBITA"] + cloud_ebita + ["20-F FY2024 表 t021 + FY2026 表 t021 [披露]"], B)
r = put(ws, r, ["  云 EBITA margin %"] + [pct(cloud_ebita[i], cloud_rev[i]) for i in range(5)] + ["[推算] ★ 3.7%→9.0% 单调改善，五年翻 2.5 倍"], GRN, None, "0.0")
r = put(ws, r, ["AIDC 国际数字商业 收入"] + aidc_rev + ["[披露]"])
r = put(ws, r, ["AIDC Adjusted EBITA"] + aidc_ebita + ["[披露] FY2026 亏损收窄 86%，AliExpress 首次经营盈利"])
r = put(ws, r, ["合并收入"] + rev + ["[披露] 双源一致"], B)
r = put(ws, r, ["  合并收入 YoY %"] + yoy(rev) + ["[推算] 增速逐年下台阶：+1.8%→+8.3%→+5.9%→+2.7%"], None, None, "0.0")
r = put(ws, r, ["合并 Adjusted EBITA"] + cons_ebita + ["[披露]"], B)
r = put(ws, r, ["  合并 EBITA YoY %"] + yoy(cons_ebita) + ["[推算] ★ FY2026 -55.8%，五年首次崩塌"], RED, None, "0.0")
r = put(ws, r, ["  合并 EBITA margin %"] + [pct(cons_ebita[i], rev[i]) for i in range(5)] + ["[推算] 15.3%→7.5%，回到 2022 年以下"], None, None, "0.0")
r += 1

r = put(ws, r, ["B. 四分部口径（FY2026 20-F，仅三年可得）", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["RMB 百万", "FY2024", "FY2025", "FY2026", "", "", "出处 / 备注"], H, F_H)
r = put(ws, r, ["中国电商集团 收入"] + cn_ec_rev + ["", "", "[披露]"])
r = put(ws, r, ["  — 客户管理收入 CMR"] + cmr + ["", "", "[披露] +5.2%（FY2026）"])
r = put(ws, r, ["  — 即时零售 quick commerce"] + quick + ["", "", "[披露] ★ +46.5%，全公司增速第二"], B)
r = put(ws, r, ["中国电商 Adjusted EBITA"] + cn_ec_ebita + ["", "", "[披露]"])
r = put(ws, r, ["  中国电商 EBITA margin %"] + [pct(cn_ec_ebita[i], cn_ec_rev[i]) for i in range(3)] + ["", "", "[推算] ★ 38.0%→19.4%，腰斩"], RED, None, "0.0")
r = put(ws, r, ["All others 收入"] + allo_rev + ["", "", "[披露] FY2026 降系业务划转，非经营性下滑"])
r = put(ws, r, ["All others EBITA（含 AI Labs/Qwen）"] + allo_ebita + ["", "", "[披露] 亏损扩大 276%"], RED)
r = put(ws, r, ["Unallocated EBITA"] + unalloc_ebita + ["", "", "[披露]"])
r = put(ws, r, ["分部间抵销 EBITA"] + elim_ebita + ["", "", "[披露]"])
r += 1

r = put(ws, r, ["C. ★★ 合并 EBITA 变动归因（FY2025→FY2026，本底稿最重要的一张表）", "", "", "", "", "", ""], B, F_K)
r = put(ws, r, ["分部", "FY2025", "FY2026", "变动额", "占总变动 %", "方向", "解读"], H, F_H)
delta_total = cons_ebita[4] - cons_ebita[3]
attrib_rows = [
    ("中国电商集团", cn_ec_ebita[1], cn_ec_ebita[2], "即时零售（淘宝闪购）补贴战——不是 AI"),
    ("All others（含 AI Labs/Qwen）", allo_ebita[1], allo_ebita[2], "模型/应用层净失血，是 AI 相关的那一块"),
    ("AIDC 国际", aidc_ebita[3], aidc_ebita[4], "正贡献：亏损收窄"),
    ("云智能集团", cloud_ebita[3], cloud_ebita[4], "正贡献：唯一利润改善主业"),
    ("Unallocated", unalloc_ebita[1], unalloc_ebita[2], "—"),
    ("分部间抵销", elim_ebita[1], elim_ebita[2], "—"),
]
for name, a, b, note in attrib_rows:
    d = b - a
    r = put(ws, r, [name, a, b, d, round(d / delta_total * 100, 1), "拖累" if d < 0 else "贡献", note],
            RED if d < 0 else GRN)
    ws.cell(row=r - 1, column=5).number_format = "0.0"
r = put(ws, r, ["合计（勾稽校验）", cons_ebita[3], cons_ebita[4], delta_total, 100.0,
                "", "★ 结论：EBITA 减少 966.49 亿中 88.7% 来自中国电商补贴战、27.1% 来自 AI 模型层；云与国际是正贡献。"
                "→ 不能把阿里 FY2026 的利润崩塌当作「AI 摊薄回报」的证据，必须先剥离补贴战"], B, F_K)
r += 1

r = put(ws, r, ["D. 新分部口径（FY2027Q1，6-K 2026-08-20，去年同期已重述）", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["RMB 百万", "2025 年同季", "2026 年本季", "YoY %", "本季 margin %", "", "出处 / 备注"], H, F_H)
q_rows = [
    ("★ AI Cloud and Compute Services 收入", 33418, 48437, "[披露] 云智能＋T-Head 合并；外部收入亦 +45%"),
    ("★ AI Cloud EBITA", 2419, 5628, "[披露] +133%，margin 7.2%→11.6%"),
    ("AI Labs and Applications 收入", 2877, 3338, "[推算去年数=3,338÷1.16] +16%"),
    ("★ AI Labs EBITA", -3224, -13861, "[披露] 亏损扩大 330%，margin -415%（Qwen App 推理成本）"),
    ("All others 收入", 28518, 28803, "[推算去年数] +1%"),
    ("All others EBITA", 687, -3343, "[披露] 转亏"),
]
for name, a, b, note in q_rows:
    is_marg = "EBITA" in name
    r = put(ws, r, [name, a, b, pct(b - a, abs(a)), None, "", note], B if name.startswith("★") else None)
    ws.cell(row=r - 1, column=4).number_format = "0.0"
r = put(ws, r, ["AI 相关产品收入（独立口径）", None, 12376, None, pct(12376, 48437),
                "", "★ [披露] 第 12 个连续季度三位数增长；占 AI Cloud 分部收入 25.6%。本体系第一个给出 AI 收入绝对金额的中国标的"], B, F_K)
ws.cell(row=r - 1, column=5).number_format = "0.0"
r += 1
r = put(ws, r, ["★ 同口径铁证（推算）", None, None, None, None, "",
                "同一公司/同一会计口径/同一报告期：模型+应用层每 1 元收入亏 4.15 元（-13,861÷3,338）；云+算力层每 1 元收入赚 0.116 元（5,628÷48,437）。"
                "消除了跨公司比较的汇率/治理/审计口径干扰 —— 这是覆盖地图核心未解题「利润是否向下迁移」的决定性证据"], B, F_K)
r = put(ws, r, ["【本表判断】", None, None, None, None, "",
                "收入端迁移已成立（云四期连续加速 +1.5%→+2.8%→+11.0%→+34.0%→+45%，AI 收入 12 季三位数增长）；"
                "利润端未成立（云 margin 仅 9–11.6% vs 电商巅峰 38%；模型层 margin -415% 且亏损加速）。[待验证 → 支持「收入迁移、利润未迁移」]"], B, F_S)

# ═════════════════════ Sheet 2_利润表五年 ═════════════════════
ws = newsheet("2_利润表五年", "利润表五年（→ 打灯 #10/#11/#12 / 分析文件第二节 利润成色）", [36, 14, 14, 14, 14, 14, 50])
r = 3
r = put(ws, r, ["RMB 百万"] + FY + ["出处 / 备注"], H, F_H)
lines = [
    ("营业额", rev, "20-F 合并利润表 [披露]；双源一致", None),
    ("毛利", gp, "[披露]", None),
    ("  毛利率 %", [pct(gp[i], rev[i]) for i in range(5)], "[推算] 稳定在 36.8%–39.9%，无恶化", None),
    ("销售及分销费用", selling, "★ [披露] FY2026 单项增 1,010 亿（+70.1%），即时零售补贴战", RED),
    ("  销售费用/收入 %", [pct(selling[i], rev[i]) for i in range(5)], "[推算] ★ 14.0%→11.9%→12.2%→14.5%→23.9%", RED),
    ("行政开支", admin, "[披露] FY2026 反而减少 111.57 亿", None),
    ("研发费用", rd, "[披露] FY2026 +93.82 亿（+16.4%）——注意 AI 投入主要走 capex 而非费用", None),
    ("  研发/收入 %", [pct(rd[i], rev[i]) for i in range(5)], "[推算] 6.5%→6.5%，占比未升", None),
    ("减值及拨备", imp, "[披露] 含商誉减值；五年从未缺席", None),
    ("★ 经营利润 Income from operations", op, "[披露] EBITDA 桥表", B),
    ("  经营利润率 %", [pct(op[i], rev[i]) for i in range(5)], "★ [推算] 8.2%→11.6%→12.0%→14.1%→4.9%，五年最低（低于 FY2022）", RED),
    ("利息与投资净收益（正=收益）", int_inv, "★ [披露] FY2026 +875.12 亿，非经营性", None),
    ("所得税费用", tax, "[披露]", None),
    ("净利润（含少数股东）", ni, "[披露]", None),
    ("★ 归母净利润", attrib, "[披露] 双源一致", B),
    ("Adjusted EBITA", cons_ebita, "[披露] 非 GAAP", None),
    ("Adjusted EBITDA", ebitda, "[披露] 非 GAAP", None),
]
for name, s, note, font in lines:
    fmt = "0.0" if "%" in name else "#,##0"
    r = put(ws, r, [name] + s + [note], font, F_K if name.startswith("★") else None, fmt)
r += 1
r = put(ws, r, ["★★ 利润成色核心对照", "", "", "", "", "", ""], B, F_K)
r = put(ws, r, ["项目", "FY2025", "FY2026", "变动 %", "", "", "解读"], H, F_H)
for name, a, b, note in [
    ("归母净利润", attrib[3], attrib[4], "表观只跌 20.4%，看起来温和"),
    ("经营利润", op[3], op[4], "★ 实际跌 64.4%——这才是主业的真相"),
    ("利息与投资净收益", int_inv[3], int_inv[4], "★ 暴增 321.6%，正好填上经营利润的坑"),
    ("经营利润 ÷ 归母净利 %", None, None, "★ FY2025 = 108.3% → FY2026 = 48.4%：净利润里过半已非经营所得"),
]:
    if a is None:
        r = put(ws, r, [name, pct(op[3], attrib[3]), pct(op[4], attrib[4]), None, "", "", note], RED, None, "0.0")
    else:
        r = put(ws, r, [name, a, b, pct(b - a, abs(a)), "", "", note], RED if b < a else None)
        ws.cell(row=r - 1, column=4).number_format = "0.0"
r += 1
r = put(ws, r, ["【本节判断】", "", "", "", "", "",
                "利润成色红旗。① 经营利润率 4.9% 为五年最低，主因是销售费用单项 +1,010 亿的补贴战，不是 AI；"
                "② 归母净利的「温和下滑」由投资收益 +875.12 亿粉饰，可持续性不明；③ 结论：给阿里定价禁用 PE 单锚。[削弱买入]"], B, F_S)

# ═════════════════════ Sheet 3_现金流照妖镜 ═════════════════════
ws = newsheet("3_现金流照妖镜", "现金流照妖镜（→ 打灯 #2/#3/#4 / 分析文件第三节）", [38, 14, 14, 14, 14, 14, 50])
r = 3
r = put(ws, r, ["RMB 百万"] + FY + ["出处 / 备注"], H, F_H)
cflines = [
    ("★ 经营现金流 OCF", ocf, "[披露] 双源一致", B),
    ("  OCF YoY %", yoy(ocf), "★ [推算] FY2026 -53.4%，两年累计 -58.3%", RED),
    ("  OCF / 营业额 %", [pct(ocf[i], rev[i]) for i in range(5)], "[推算] 16.7%→23.0%→19.4%→16.4%→7.4%", RED),
    ("  OCF / 归母净利（倍）", [round(ocf[i] / attrib[i], 2) for i in range(5)], "★ [推算] 2.29→2.74→2.28→1.26→0.74：FY2026 首次 <1，利润不再变成钱", RED),
    ("购建固定资产 capex（现金流量表口径）", capex, "[披露] = Item 5.B 口径，含办公园区", None),
    ("  capex / 营业额 %", [pct(capex[i], rev[i]) for i in range(5)], "[推算] 6.2%→4.0%→3.4%→8.6%→12.3%（FY2027Q1 单季 25.2%）", None),
    ("  capex / OCF %", [pct(capex[i], ocf[i]) for i in range(5)], "★ [推算] 37.3%→17.2%→17.6%→52.6%→165.4%：投入已是造血能力的 1.65 倍", RED),
    ("★ 自由现金流 FCF（公司口径）", fcf, "[披露] 非 GAAP 调节表；capex 用剔除办公园区口径", B),
    ("折旧及摊销（现金流量表全口径）", da_cf, "[披露]", None),
    ("折旧＋房产设备减值＋土地租赁（EBITDA 桥）", dep_ppe, "[披露] 与上一行口径不同，勿混用", None),
    ("  capex / 折旧（倍）", [round(capex[i] / dep_ppe[i], 2) for i in range(5)], "★ [推算] 1.92→1.23→1.20→2.94→3.40：未来折旧压力已锁定", RED),
    ("投资活动现金净额", icf, "[披露]", None),
    ("筹资活动现金净额", fcf_fin, "[披露]", None),
]
for name, s, note, font in cflines:
    fmt = "0.00" if "倍" in name else ("0.0" if "%" in name else "#,##0")
    r = put(ws, r, [name] + s + [note], font, F_K if name.startswith("★") else None, fmt)
r += 1
r = put(ws, r, ["★★ 融资方向三段式转向（说-做对账最硬的证据）", "", "", "", "", "", ""], B, F_K)
r = put(ws, r, ["RMB 百万"] + FY + ["出处 / 备注"], H, F_H)
for name, s, note in [
    ("回购股份（现金流出）", buyback, "★ [披露] FY2026 从 866.62 亿断崖降至 76.38 亿（-91.2%）"),
    ("已付股息", dividend, "[披露] 仍在增加"),
    ("新增借款", new_debt, "[披露] FY2026 84,228，其中 71,774「主要用于资本开支与营运资金」"),
    ("偿还借款", repay, "[披露]"),
    ("发行债券（可转债＋可交换债净额）", bond_iss, "★ [披露] FY2025 起从零到 716.44 亿、333.62 亿"),
]:
    r = put(ws, r, [name] + s + [note], B if name.startswith("回购") else None)
r = put(ws, r, ["配股（FY2027Q2，2026-08-26 完成）", None, None, None, None, "HK$800 亿",
                "★ [披露] 7.10 亿股 @ HK$112.70，≈US$102 亿；募资 100% 投 AI 基建（60% 全球算力／40% 超大规模 AI 数据中心）"], B, F_K)
r += 1
r = put(ws, r, ["【本节判断】", "", "", "", "", "",
                "现金流红旗，且比利润表更严重。① OCF/归母净利 FY2026 首次跌破 1（0.74）；② capex/OCF 达 165.4%，投入超过造血；"
                "③ FCF 首次转负 -466.09 亿，FY2027Q1 单季即 -446.70 亿（≈上一整年的 96%）；"
                "④ 三年内从「年回购 866 亿」走到「单次增发 800 亿港元」，资金来源由自身造血切换为债权人＋股东出资。"
                "⑤ 与管理层「AI 已进入全面商业化」的定性直接冲突。[削弱买入]"], B, F_S)

# ═════════════════════ Sheet 4_资产负债偿债 ═════════════════════
ws = newsheet("4_资产负债偿债", "资产负债与偿债（→ 打灯 #5-#9 / 分析文件第四节）", [38, 14, 14, 14, 14, 14, 50])
r = 3
r = put(ws, r, ["RMB 百万"] + FY + ["出处 / 备注"], H, F_H)
ib = [st_loan[i] + lt_loan[i] + conv[i] for i in range(5)]
cst = [cash[i] + sti[i] for i in range(5)]
netcash = [cst[i] - ib[i] for i in range(5)]
bslines = [
    ("总资产", ta, "[披露] 双源一致", None),
    ("总负债", tl, "[披露]", None),
    ("  资产负债率 %", [pct(tl[i], ta[i]) for i in range(5)], "[推算] 36.2%→41.0%，缓升但仍健康", None),
    ("归母股东权益", eq_p, "[披露]", None),
    ("少数股东权益", eq_m, "[披露] 五年降 54%（回购子公司少数股权）", None),
    ("★ 物业厂房及设备 PP&E", ppe, "★ [披露] FY2026 +39.0%，算力资产集中上表", B),
    ("  PP&E / 总资产 %", [pct(ppe[i], ta[i]) for i in range(5)], "[推算] 10.1%→14.8%：资产结构正在变重", None),
    ("商誉", goodwill, "[披露] 五年缓降（持续计提减值）；FY2026 占总资产 13.0%", None),
    ("现金及等价物", cash, "[披露]", None),
    ("短期投资", sti, "[披露]", None),
    ("★ 现金＋短期投资", cst, "[推算]", B),
    ("短期贷款", st_loan, "[披露]", None),
    ("长期贷款", lt_loan, "[披露]", None),
    ("可转换票据及债券", conv, "[披露] FY2025 起从零到 668.37 亿", None),
    ("★ 有息负债合计", ib, "[推算] 短贷＋长贷＋可转债/可交换债", B),
    ("★★ 净现金（现金＋短投 − 有息负债）", netcash, "★ [推算] FY2024 峰值 5,026 亿 → FY2026 1,744 亿", RED),
    ("净流动资产", nwc, "★ [披露] 3,314 亿→1,344 亿，流动性缓冲被消耗", RED),
    ("流动资产合计", ca, "[披露]", None),
    ("流动负债合计", cl, "[披露]", None),
    ("  流动比率（倍）", [round(ca[i] / cl[i], 2) for i in range(5)], "[推算] 1.66→1.28，仍 >1 但单调下降", None),
]
for name, s, note, font in bslines:
    fmt = "0.00" if "倍" in name else ("0.0" if "%" in name else "#,##0")
    r = put(ws, r, [name] + s + [note], font, F_K if name.startswith("★") else None, fmt)
r += 1
r = put(ws, r, ["★★ 净现金消耗速度（AI 投入的真实代价）", "", "", "", "", "", ""], B, F_K)
r = put(ws, r, ["项目", "FY2024（峰值）", "FY2026", "变动额", "变动 %", "", "解读"], H, F_H)
r = put(ws, r, ["净现金", netcash[2], netcash[4], netcash[4] - netcash[2],
                pct(netcash[4] - netcash[2], netcash[2]), "",
                "★ 两年蒸发 3,282 亿（-65.3%）。同期回购只花了 943 亿、分红 648 亿——差额去了 capex（两年 2,120 亿）与营运资金"], RED)
ws.cell(row=r - 1, column=5).number_format = "0.0"
r = put(ws, r, ["公司口径「现金及流动投资」FY2027Q1", None, 474505, None, None, "",
                "⚠️ 6-K 口径含 other treasury investments，与本表东财口径（FY2026 316,894）差约 1,576 亿。口径差异待核，引用时须声明用哪个"], None, F_S)
r += 1
r = put(ws, r, ["【本节判断】", "", "", "", "", "",
                "资产负债表仍是全场最强项，但趋势在恶化。① 净现金 1,744 亿仍厚，按 FY2027Q1 单季 FCF -446.7 亿的速度可撑约 4 个季度而不需再融资"
                "（这解释了为何 8 月就配股）；② 资产结构由轻转重（PP&E 占比 10.1%→14.8%），未来折旧刚性上升；"
                "③ 资产负债率 41% 无偿债风险。[中性偏削弱：不是偿债问题，是「钱够不够撑到 AI 变现」的问题]"], B, F_S)

# ═════════════════════ Sheet 5_估值对比 ═════════════════════
ws = newsheet("5_估值对比", "估值：同行横向对比（取价日 2026-09-14，东财源）", [22, 14, 14, 16, 10, 10, 12, 44])
r = 3
r = put(ws, r, ["公司", "代码", "收盘 HKD", "总市值 亿HKD", "PB", "ROE %", "股息率 %", "备注"], H, F_H)
peers = [
    ("★ 阿里巴巴-W", "09988", 105.80, 21038, 1.67, 1.0, 0.94, "本标的；当日 -1.58%"),
    ("腾讯控股", "00700", 430.60, 39198, 2.99, 9.97, 1.23, "对照：平台＋自研模型，但 AI capex 强度远低"),
    ("百度集团-W", "09888", 89.60, 2449, 0.78, 2.14, None, "对照：同为云＋模型双线，PB 更低（市场更不信）"),
    ("金山云", "03896", 5.085, 230, 2.25, -4.8, None, "★ 唯一纯云标的：PB 2.25 > 阿里 1.67 而 ROE -4.8%"),
    ("京东集团-SW", "09618", 106.10, 2990, 1.19, 5.51, 3.66, "对照：电商同业，无 AI 云叙事"),
    ("美团-W", "03690", 74.75, 4616, 2.37, -2.92, None, "对照：即时零售补贴战的另一方，同样被打亏"),
]
for name, code, px, mc, pb, roe, dy, note in peers:
    r = put(ws, r, [name, code, px, mc, pb, roe, dy, note], B if name.startswith("★") else None)
    for col, f in ((3, "0.000"), (5, "0.00"), (6, "0.00"), (7, "0.00")):
        ws.cell(row=r - 1, column=col).number_format = f
r += 1
r = put(ws, r, ["★ 纯云对比的含义", "", "", "", "", "", "",
                "市场愿意给亏钱的纯云（金山云 PB 2.25、ROE -4.8%）比阿里（PB 1.67）更高的 PB。"
                "→ 阿里集团 PB 的主锤仍是电商资产，云的估值溢价未被充分定价。含义双向：既是「云没被定价」的重估空间，"
                "也是「市场不认为阿里的云能独立变现」的怀疑票。分拆预期是潜在触发器 [待验证]"], B, F_S)
r = put(ws, r, ["⚠️ 口径警告 1", "", "", "", "", "", "",
                "东财 ROE 1.0% 为其自有口径（似最新报告期年化），与 AKShare 年度 ROE_AVG 9.91%（FY2026）不同口径，不得横向比"], None, None)
r = put(ws, r, ["⚠️ 口径警告 2", "", "", "", "", "", "",
                "东财股本显示 198.85 亿股，20-F 口径 2026-05-18 为 186.70 亿股、配股后约 193.80 亿股，差约 5 亿股（2.6%）。市值/PB 引用须声明口径 [待核]"], None, None)
r = put(ws, r, ["⚠️ 可比性声明", "", "", "", "", "", "",
                "六家同处 AI 投入期但商业模式不同（平台/纯云/电商/本地生活），横向 PB 只能做粗量级参照，不构成买点。"
                "T 类型尺子以纵向分位＋分部 SOTP 为主（见 5b）"], None, None)

# ═════════════════════ Sheet 5b_股价与估值 ═════════════════════
ws = newsheet("5b_股价与估值", "股价与估值（Step 5b，取价日 2026-09-14；日线降级用腾讯源不复权）", [24, 13, 13, 13, 13, 13, 13, 46])
r = 3
r = put(ws, r, ["A. 历年股价（腾讯源不复权；涨跌幅为前复权）", "", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["年度", "年末收盘", "年内均价", "年内最高", "年内最低", "涨跌 %", "", "备注"], H, F_H)
for y, cl_, av, hi, lo, ch, note in [
    ("2019（仅24日）", 207.20, 200.44, 212.80, 187.60, None, "2019-11-26 港股二次上市，序列起点"),
    ("2020", 232.60, 233.15, 307.40, 170.00, 12.6, "★ 307.40 为历史最高"),
    ("2021", 118.90, 193.40, 267.00, 109.90, -50.1, "反垄断＋教培监管冲击"),
    ("2022", 86.25, 95.00, 133.00, 61.45, -28.8, "★ 61.45 为历史最低"),
    ("2023", 75.60, 87.50, 117.50, 68.80, -12.0, "分拆计划反复"),
    ("2024", 82.40, 79.46, 115.10, 65.45, 11.9, "回购高峰年"),
    ("2025", 142.80, 128.61, 185.10, 78.10, 78.5, "★ AI 叙事驱动，年涨 78.5%"),
    ("2026（至09-11）", 105.80, 128.56, 173.50, 89.50, -25.4, "★ 高点 173.50（01-28），当前回撤 -39.0%"),
]:
    r = put(ws, r, [y, cl_, av, hi, lo, ch, "", note], B if y.startswith("2026") else None, None, "0.00")
r += 1
r = put(ws, r, ["B. 每股指标与汇率（BVPS/EPS 取 AKShare 财务指标，汇率为估算）", "", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["项目"] + FY + ["", "出处 / 口径"], H, F_H)
r = put(ws, r, ["总股本（亿股，期末）", None, None, None, 186.70, 193.80, "", "[披露] 20-F 封面 2026-05-18 = 186.70 亿股；配股 7.10 亿股后 ≈193.80 亿"], None, None, "0.00")
r = put(ws, r, ["每股净资产 BVPS（CNY）", 48.184, 50.265, 50.152, 51.374, 53.746, "", "[东财指标] BPS"], None, None, "0.000")
r = put(ws, r, ["每股基本盈利 EPS（CNY）", 2.87, 3.46, 3.95, 6.89, 5.70, "", "[东财指标] BASIC_EPS"], None, None, "0.00")
r = put(ws, r, ["★ 每股经营利润（CNY，推算）", None, None, None, None, round(50150 / 1938, 3), "", "★ [推算] 经营利润 50,150 ÷ 193.80 亿股 = 2.588 元——这是给阿里定价该用的分母"], B, F_K, "0.000")
r = put(ws, r, ["汇率 CNY→HKD（3-31）", 1.235, 1.142, 1.084, 1.072, 1.106, "", "⚠️ [估算] 非权威源，PB/PE 绝对值只作量级参照"], None, F_S, "0.000")
r += 1
r = put(ws, r, ["C. 当前估值与历史分位（窗口 2019-11-26→2026-09-11，1,674 交易日 ≈ 6.8 年 = 全历史，可写「近 7 年」）", "", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["指标", "当前值", "数值分位 %", "便宜程度分位 %", "中位", "区间下限", "区间上限", "口径与备注"], H, F_H)
for name, cur, q, cheap, med, lo, hi, note in [
    ("★ 股价（无汇率依赖）", 105.80, 42.4, 57.6, 116.10, 61.45, 307.40, "★ 最可信的锚（FX-free）。挡位：10%=76.00｜25%=85.75｜50%=116.10｜75%=180.48｜90%=228.34"),
    ("PB（脚本 FY2025 每股数）", 1.952, 44.7, 55.3, 2.093, 1.088, 10.468, "[推算] 挡位对应股价：PB1.30→71.69／PB1.47→80.86／PB2.09→115.27"),
    ("PB（东财快照，FY2026）", 1.67, None, None, None, None, None, "以东财为准（东财自行处理汇率）；本底稿手算 FY2026 口径为 1.78"),
    ("PE 静态（归母净利口径）", 14.55, 16.9, 83.1, 23.015, 7.45, 61.97, "⚠️ 脚本用 FY2025 EPS；按 FY2026 EPS 5.70 手算 = 16.8 倍"),
    ("★★ PE（经营利润口径）", 37.0, None, None, 23.015, None, None, "★★ [推算] 105.80 ÷（2.588×1.106）。相对历史中位 23.0 高出 61% —— 与静态 PE 结论完全相反"),
    ("股息率 %", 0.94, None, None, None, None, None, "[东财快照] 极低，不构成支撑锚"),
]:
    r = put(ws, r, [name, cur, q, cheap, med, lo, hi, note],
            B if name.startswith("★") else None, None, "0.000")
    if name.startswith("★★"):
        for c in range(1, 9):
            ws.cell(row=r - 1, column=c).fill = F_K
r += 1
r = put(ws, r, ["★★ D. 估值的一刀：「便宜」是利润口径造出来的", "", "", "", "", "", "", ""], B, F_K)
r = put(ws, r, ["口径", "PE 倍数", "相对历史中位 23.0", "", "", "", "", "推算链"], H, F_H)
r = put(ws, r, ["表观静态 PE（归母净利 103,592）", 16.8, "-27%（历史 17% 分位）", "", "", "", "", "看上去很便宜"], None, None, "0.0")
r = put(ws, r, ["经营利润口径 PE（50,150）", 37.0, "+61%", "", "", "", "", "★ 实际不便宜"], RED, F_K, "0.0")
r = put(ws, r, ["差额来源", 87512, "利息与投资净收益", "", "", "", "", "归母净利 1,035.92 亿中 875.12 亿来自投资收益（与 AI 业务无关、不可持续）。剔除后市场在为经营盈利付 37 倍"], B)
r = put(ws, r, ["补充锚 EV/Adj EBITA", 18.7, "[推算，未扣有息负债]", "", "", "", "", "市值 21,038 亿HKD÷1.106=19,022 亿CNY − 现金及流动投资 4,745 亿 = EV 14,277 亿 ÷ EBITA 764.16 亿。不减净现金为 24.9 倍；未扣有息负债 1,425 亿，真 EV 更高，本值为下限"], None, None, "0.0")
r += 1
r = put(ws, r, ["E. 配股折价率与市场反应（机构定价锚的检验）", "", "", "", "", "", "", ""], B, F_S)
r = put(ws, r, ["日期", "开", "收", "高", "低", "成交量 亿股", "", "事件"], H, F_H)
for d, o, c_, h_, l_, v, e in [
    ("2026-08-20", 126.6, 126.2, 129.0, 124.3, 0.92, "FY2027Q1 业绩发布"),
    ("2026-08-21", 129.2, 123.0, 129.8, 121.5, 1.47, "★ 配股前最后一个交易日（折价率基准）"),
    ("2026-08-24", 113.1, 112.5, 113.5, 110.1, 3.63, "★ 配股公告后首日：跌 8.5%，成交量放大至平日 2.5 倍"),
    ("2026-08-26", 115.7, 116.6, 117.8, 115.6, 1.34, "配股完成"),
]:
    r = put(ws, r, [d, o, c_, h_, l_, v, "", e], None, None, "0.00")
r = put(ws, r, ["★ 配股折价率", round(112.70 / 123.00 - 1, 4), "", "", "", "", "", "★ [推算] 112.70÷123.00−1 = -8.37%，港股大额配售典型区间，未出现急切式深折 → 需求尚在"], B, F_K, "0.00%")
r = put(ws, r, ["★★ 配股价已破发", round(105.80 / 112.70 - 1, 4), "", "", "", "", "", "★★ [推算] 现价 105.80 < 配售价 112.70，配股投资人 18 个交易日浮亏 -6.1%。含义：HK$112.70 是一个已被市场否定的机构锚，不能当「专业投资人背书的底」用"], RED, F_K, "0.00%")
r = put(ws, r, ["摊薄比例", round(710 / 18669.888147 * 100, 2), "%", "", "", "", "", "[推算] 7.10 亿股 ÷ 186.70 亿股（2026-05-18 已发行）= 3.80%"], None, None, "0.00")
r += 1
r = put(ws, r, ["【本表判断】", "", "", "", "", "", "",
                "① 唯一可信的纵向锚是股价分位：当前 105.80 处近 7 年 42.4% 数值分位（便宜程度 57.6%）—— 中位偏下，不是深度低估；"
                "② 表观 PE 17% 分位是口径幻觉，经营利润口径 PE 高于历史中位 61%；"
                "③ 机构配售价 112.70 已破发，不构成价格底；④ 股息率 0.94% 无支撑。"
                "结论：现价不提供安全边际。[削弱买入]"], B, F_S)
r = put(ws, r, ["⚠️ 待补", "", "", "", "", "", "",
                "BABA ADS 价与人民币柜台 89988 未取（脚本不支持美股）。不阻塞：双重主要上市，1 ADS = 8 股，两地价差历史极小"], None, None)

# ═════════════════════ Sheet 6_取数日志 ═════════════════════
ws = newsheet("6_取数日志", "取数日志（查不到出处的数字 = 没有的数字）", [5, 34, 24, 34, 13, 46])
r = 3
r = put(ws, r, ["#", "数据项", "来源", "取数方式 / 定位", "取数日期", "备注与异常"], H, F_H)
logs = [
    ("披露文件清单", "SEC EDGAR submissions API", "data.sec.gov/submissions/CIK0001577552.json", "2026-09-14", "确认 CIK 0001577552，20-F 十二份；须带 User-Agent，macOS 需 certifi CA"),
    ("20-F FY2022–FY2026 五份原文", "SEC EDGAR Archives", "filing index.json → 主文档 HTML（7.7–14.1MB）", "2026-09-14", "落 1.原始材料/a_年报/"),
    ("⚠️ FY2025 正文补取", "SEC EDGAR Archives", "Exhibit 15.5", "2026-09-14", "★ FY2025 主文档只有财务报表，MD&A/业务描述放在 Exhibit 15.5。首次提取仅 0.51M 字符、'Cloud Intelligence' 仅 9 次 → 已用 edgar_fix_fy2025.py 补取"),
    ("6-K EX-99.1 五份", "SEC EDGAR 6-K exhibit", "按月份窗口筛 index.json → 下载 ex99", "2026-09-14", "June Quarter 2026 业绩／配股 launch＋closing／CMC List／June Quarter 2025 业绩"),
    ("年度分部收入与 EBITA", "20-F FY2024 + FY2026", "extract_tables.py：正则切 <table> → pandas.read_html", "2026-09-14", "★ iXBRL 表格不能用剥标签方式读（会被拆成空单元格行）。落 tables/*_REV.csv、*_EBITA.csv"),
    ("OCF/FCF 五年序列", "20-F FCF 非 GAAP 调节表", "extract_cf.py 关键词 'Free cash flow'", "2026-09-14", "tables/FY2024_t025_FCF_0.csv、FY2026_t025_FCF_0.csv"),
    ("折旧与 EBITDA 桥", "20-F EBITDA 调节表", "extract_cf.py 关键词 'Depreciation'", "2026-09-14", "tables/FY2024_t022_DEP_0.csv、FY2026_t022_DEP_0.csv"),
    ("capex 与资本承诺", "20-F Item 5.B", "grep 定位", "2026-09-14", "★ 两个 capex 口径差 4,042（含/不含办公园区），已在 Sheet0 声明"),
    ("三大报表五年骨架", "AKShare 东财港股接口", "fetch_akshare_data.py → akshare_09988/*.csv", "2026-09-14", "441/261/370/9 行；⚠️ CURRENCY 字段错标 HKD（IS_CNY_CODE=0），实际人民币"),
    ("交叉验证（Step 4）", "20-F vs 东财", "锚点三科目对比", "2026-09-14", "★ 营业额 1,023,670／归母净利 103,592／总资产 1,909,570 三项误差 0%，双源一致"),
    ("同行快照 PB/ROE/股息率", "fetch_stock_price.py snapshot（东财）", "6 只：09988/00700/09888/03896/09618/03690", "2026-09-14", "落 price/snapshot_peers.csv；东财 ROE 为自有口径，勿与年度 ROE 混比"),
    ("★ 日线与年度序列（降级）", "fetch_stock_price.py --source tencent", "daily/annual 子命令，不复权", "2026-09-14", "★ 东财 RemoteDisconnected 限流 → 降级腾讯源。同一张表混用两源已逐行标注（Sheet5b）"),
    ("PB/PE 历史分位", "fetch_stock_price.py valuation", "--bvps/--eps/--fx，--lag-months 2", "2026-09-14", "★ 脚本按自然年末＋lag 生效，对 3-31 财年失效，实际用了 FY2025 每股数 → FY2026 口径已手工补算"),
    ("汇率 CNY→HKD", "非权威源", "各年 3-31 近似值", "2026-09-14", "⚠️ [估算]。PB/PE 绝对值与分位仅作量级参照，定价优先用 FX-free 的股价分位"),
    ("配股条款", "6-K 2026-08-24 / 08-26 EX-99.1", "全文精读", "2026-09-14", "7.10 亿股 @ HK$112.70，HK$800 亿；Regulation S 未在美注册；募资 60% 全球算力／40% 超大规模 AI 数据中心"),
    ("CMC List 事件", "6-K 2026-06-09 EX-99.1", "全文精读", "2026-09-14", "法律效果仅限美国国防部不得采购；不附加出口管制或制裁。记为尾部风险（小米 2021 先例可诉可除名）"),
    ("BIS 云管制风险", "20-F FY2026 风险因素", "grep 'ICTS'", "2026-09-14", "拟制定云计算 ICTS 规则，原文称可能限制云智能集团进入美国市场——与配股 60% 用途正面相撞"),
    ("市占与第三方口径", "20-F 引用第三方", "Gartner/IDC/Omdia/Frost&Sullivan", "2026-09-14", "[第三方，公司引用] 全球第四大 IaaS；中国最大公有云；Qwen 企业级调用 32% 第一；中国 AI 云 38.1% 第一"),
    ("⚠️ 巨潮通道失败", "巨潮资讯网", "download_cninfo_reports.py --keyword 年報/年度報告/年报", "2026-09-14", "★ 检索 0 条（简体「年报」检索到 5 条但下载 0 份）→ 全量转 EDGAR。港股年报走巨潮不可靠，已记入经验"),
    ("⚠️ 未取到（待补）", "—", "—", "—", "① FY2026 中间三季度业绩（Sept'25/Dec'25/Mar'26）→ 云分部逐季序列只有两点 ② BABA ADS 与 89988 价 ③ 东财股本口径差 5 亿股 ④ 分部 capex 归属（决定「云自身造血率」精度） ⑤ 经营租赁承诺与在建工程明细"),
]
for i, (item, src, how, when, note) in enumerate(logs, start=1):
    r = put(ws, r, [i, item, src, how, when, note])
    ws.cell(row=r - 1, column=6).alignment = WRAP
    ws.cell(row=r - 1, column=2).alignment = WRAP
    if item.startswith(("★", "⚠️")):
        for c in range(1, 7):
            ws.cell(row=r - 1, column=c).fill = F_K

del wb["Sheet"]
wb.save(OUT)
print(f"已生成：{OUT}")
print("Sheets:", wb.sheetnames)
