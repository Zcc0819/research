import warnings, logging, re
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
f="../1.原始材料/首次公开发行股票招股说明书.pdf"
PS=[2,28,40,42,44,49,50,203,442]
with pdfplumber.open(f) as pdf:
    for p in PS:
        t=pdf.pages[p-1].extract_text() or ""
        print(f"\n############ p.{p}"); print(t[:2600])
