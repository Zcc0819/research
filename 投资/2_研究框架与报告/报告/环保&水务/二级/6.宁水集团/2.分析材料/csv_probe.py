import csv
base="/Users/zhangarmanda/PycharmProjects/Personal/投资/2_研究框架与报告/报告/环保&水务/二级/6.宁水集团/2.分析材料/"

def load(fn):
    with open(base+fn, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

WANT = {
 "A_利润表_年度.csv": ["REPORT_DATE","TOTAL_OPERATE_INCOME","OPERATE_INCOME","OPERATE_COST","TOTAL_OPERATE_COST",
    "SALE_EXPENSE","MANAGE_EXPENSE","RESEARCH_EXPENSE","FINANCE_EXPENSE","INVEST_INCOME","INVEST_JOINT_INCOME",
    "TOTAL_PROFIT","NETPROFIT","PARENT_NETPROFIT","DEDUCT_PARENT_NETPROFIT","MINORITY_INTEREST"],
 "A_现金流量表_年度.csv": ["REPORT_DATE","NETCASH_OPERATE","NETCASH_INVEST","NETCASH_FINANCE",
    "CONSTRUCT_LONG_ASSET","ASSIGN_DIVIDEND_PORFIT","END_CASH_EQUIVALENTS","BEGIN_CASH_EQUIVALENTS","NETPROFIT"],
 "A_资产负债表_年度.csv": ["REPORT_DATE","TOTAL_ASSETS","TOTAL_PARENT_EQUITY","TOTAL_EQUITY","MONETARYFUNDS",
    "SHORT_LOAN","LONG_LOAN","NONCURRENT_LIAB_1YEAR","ACCOUNTS_RECE","NOTE_RECE","NOTE_ACCOUNTS_RECE",
    "FIN_RECE","CONTRACT_ASSET","INVENTORY","FIXED_ASSET","TOTAL_LIABILITIES","LEASE_LIAB","BOND_PAYABLE"],
}
for fn, keys in WANT.items():
    rows=load(fn)
    print("=====", fn)
    avail=[k for k in keys if k in rows[0]]
    miss=[k for k in keys if k not in rows[0]]
    print("MISSING:", miss)
    print("\t".join(avail))
    for r in rows:
        if r["REPORT_DATE"][:4] in ("2019","2020","2021","2022","2023","2024","2025"):
            out=[]
            for k in avail:
                v=r[k]
                if k=="REPORT_DATE": out.append(v[:10])
                else:
                    try: out.append("%.1f"%(float(v)/1e6))
                    except: out.append("-")
            print("\t".join(out))
    print()
