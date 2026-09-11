import pdfplumber, os, re

BASE = "/Users/zhangarmanda/PycharmProjects/Personal/投资/2_研究框架与报告/报告/环保&水务/二级/6.宁水集团/1.原始材料"
JOB = {
    "2020年年度报告.pdf": (99, 103),
    "2021年年度报告.pdf": (100, 105),
    "2022年年度报告.pdf": (103, 105),
    "2023年年度报告.pdf": (115, 118),
    "2024年年度报告.pdf": (131, 133),
    "2025年年度报告.pdf": (134, 136),
}
PAT = [re.compile(r"（[一二三四五六]）\s*\S{2,20}"),
       re.compile(r"账面价值为?[\d,\.]+万?元"),
       re.compile(r"占营业收入比例为?[\d\.]+%"),
       re.compile(r"合并营业收入为?[\d,\.]+\s*万?元"),
       re.compile(r"坏账准备[^。]{0,40}"),
       re.compile(r"存货[^。]{0,30}账面价值[^。]{0,30}")]
for f, (a, b) in JOB.items():
    print("=" * 10, f)
    with pdfplumber.open(os.path.join(BASE, f)) as pdf:
        buf = ""
        for p in range(a, b + 1):
            buf += (pdf.pages[p - 1].extract_text() or "").replace("\n", "").replace(" ", "")
        for pt in PAT:
            for m in dict.fromkeys(pt.findall(buf)):
                print("   ", m)
