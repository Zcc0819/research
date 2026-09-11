import warnings, logging, re
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
f="../1.原始材料/c_临时公告/宁波水表（集团）股份有限公司2024年股票期权激励计划（草案）.pdf"
with pdfplumber.open(f) as pdf:
    print("总页", len(pdf.pages))
    for i,pg in enumerate(pdf.pages):
        t=pg.extract_text() or ""
        if re.search(r"行权条件|业绩考核|归属期|营业收入|净利润|考核年度|第一个行权期|公司层面", t):
            keep=[l for l in t.split("\n") if re.search(r"行权|考核|营业收入|净利润|归属|增长率|%|20\d\d年", l)]
            if keep: print(f"\n--- p.{i+1}"); print("\n".join(keep[:24]))
