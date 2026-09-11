import warnings, logging
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
PAGES={"2020":[24,25],"2021":[29,30],"2022":[30],"2023":[38,39],"2024":[46,47]}
for y,ps in PAGES.items():
    with pdfplumber.open(f"../1.原始材料/{y}年年度报告.pdf") as pdf:
        for p in ps:
            for tb in pdf.pages[p-1].extract_tables():
                rows=["|".join((c or "").replace("\n","") for c in r) for r in tb]
                body="\n".join(rows)
                if "库存量" in body and ("智能水表" in body or "智能水" in body):
                    print(f"\n### FY{y} p.{p}"); print(body)
