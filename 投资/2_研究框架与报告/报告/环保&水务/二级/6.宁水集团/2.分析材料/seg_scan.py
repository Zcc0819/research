import warnings, logging, re
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
PAGES={"2020":[23,24,26],"2021":[28,29,31],"2022":[29,30,31,32],"2023":[37,38,40],"2024":[45,46,48]}
for y,ps in PAGES.items():
    print("\n########## FY"+y)
    with pdfplumber.open(f"../1.原始材料/{y}年年度报告.pdf") as pdf:
        for p in ps:
            t=pdf.pages[p-1].extract_text() or ""
            keep=[l for l in t.split("\n") if re.search(r"智能水表|机械水表|水表配件|智慧水务|软件|仪器仪表|内销|外销|境内|境外|直销|经销|毛利率|前五名客户|前五名供应商|占年度销售总额|占年度采购总额|库存量|生产量|销售量", l)]
            if keep: print(f"--- p.{p}"); print("\n".join(keep[:30]))
