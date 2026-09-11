import warnings, logging, re
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
PAGES={"2020":[99,100,101,102],"2021":[100,101,102,103,104],"2022":[103,104],"2023":[115,116,117],"2024":[131,132],"2025":[134,135]}
for y,ps in PAGES.items():
    print("\n########## FY"+y)
    with pdfplumber.open(f"../1.原始材料/{y}年年度报告.pdf") as pdf:
        for p in ps:
            t=pdf.pages[p-1].extract_text() or ""
            keep=[l for l in t.split("\n") if re.search(r"关键审计事项|^（[一二三四五]）|事项[一二三四五]|审计意见|应收账款|收入确认|存货|商誉|账面价值|占营业收入|占总资产|万元|无保留", l)]
            if keep: print(f"--- p.{p}"); print("\n".join(keep[:22]))
