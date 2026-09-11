import warnings, logging, re, sys
warnings.filterwarnings("ignore"); logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber
f="../1.原始材料/首次公开发行股票招股说明书.pdf"
KW={"issue":["本次发行概况","发行股数","发行价格","市盈率"],
    "cust":["前五名客户","前五大客户","向前五名供应商"],
    "fin":["主要财务数据及财务指标","加权平均净资产收益率"],
    "use":["募集资金投资项目","募集资金运用","项目投资总额"],
    "risk":["风险因素","毛利率下降的风险","应收账款"],
    "share":["实际控制人","控股股东"]}
with pdfplumber.open(f) as pdf:
    print("总页",len(pdf.pages))
    hits={}
    for i,pg in enumerate(pdf.pages):
        t=pg.extract_text() or ""
        for k,kws in KW.items():
            if any(w in t for w in kws): hits.setdefault(k,[]).append(i+1)
    for k in KW: print(k, hits.get(k,[])[:25])
