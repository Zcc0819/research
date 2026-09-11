import warnings, logging, os
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
os.makedirs("pdf_extracted", exist_ok=True)
PAGES={"2020":[3,4,18,20,34,39],"2021":[2,10,39,44],"2022":[2,10,39,50],"2023":[13,48,61],"2024":[16,57,70]}
for y,ps in PAGES.items():
    with pdfplumber.open(f"../1.原始材料/{y}年年度报告.pdf") as pdf:
        for p in ps:
            t=pdf.pages[p-1].extract_text() or ""
            open(f"pdf_extracted/FY{y}_indrisk_p{p}.txt","w").write(t)
            print(f"\n############ FY{y} p.{p}"); print(t[:1900])
