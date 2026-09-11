import warnings, logging
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
with pdfplumber.open("../1.原始材料/2020年年度报告.pdf") as pdf:
    for i in range(0,8):
        t=pdf.pages[i].extract_text() or ""
        if "利润分配预案" in t or "每10股派" in t:
            print(f"--- p.{i+1}"); print(t[:1200]); break
