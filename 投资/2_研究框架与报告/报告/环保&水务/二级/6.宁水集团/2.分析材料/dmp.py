import pdfplumber, os, sys

BASE = "/Users/zhangarmanda/PycharmProjects/Personal/投资/2_研究框架与报告/报告/环保&水务/二级/6.宁水集团/1.原始材料"
# usage: python3 dmp.py <file> p1 p2 p3 ...
f = sys.argv[1]
pages = [int(x) for x in sys.argv[2:]]
with pdfplumber.open(os.path.join(BASE, f)) as pdf:
    for p in pages:
        t = pdf.pages[p - 1].extract_text() or ""
        print("\n" + "=" * 12 + " %s p.%d " % (f, p) + "=" * 12)
        print(t)
