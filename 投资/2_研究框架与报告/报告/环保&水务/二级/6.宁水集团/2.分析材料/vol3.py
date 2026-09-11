import warnings, logging
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
with pdfplumber.open("../1.原始材料/2024年年度报告.pdf") as pdf:
    print(pdf.pages[46].extract_text()[:900])
