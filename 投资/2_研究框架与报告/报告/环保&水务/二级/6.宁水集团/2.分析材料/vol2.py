import warnings, logging, re
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
for y,ps in {"2021":[29,30],"2023":[38,39],"2024":[46,47]}.items():
    with pdfplumber.open(f"../1.原始材料/{y}年年度报告.pdf") as pdf:
        for p in ps:
            t=pdf.pages[p-1].extract_text() or ""
            if "库存量" in t:
                idx=t.find("库存量")
                print(f"\n### FY{y} p.{p}"); print(t[max(0,idx-300):idx+700])
