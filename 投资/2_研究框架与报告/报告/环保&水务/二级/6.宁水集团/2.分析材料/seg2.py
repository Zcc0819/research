import warnings, logging
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
PAGES={"2021":[28,29],"2022":[29,30],"2023":[37,38],"2024":[45,46]}
for y,ps in PAGES.items():
    print("\n########## FY"+y)
    with pdfplumber.open(f"../1.原始材料/{y}年年度报告.pdf") as pdf:
        for p in ps:
            for ti,tb in enumerate(pdf.pages[p-1].extract_tables()):
                rows=["|".join((c or "").replace("\n","") for c in r) for r in tb]
                body="\n".join(rows)
                if any(k in body for k in ["智能水表","机械水表","仪器仪表"]):
                    print(f"--- p.{p} table{ti}"); print(body)
