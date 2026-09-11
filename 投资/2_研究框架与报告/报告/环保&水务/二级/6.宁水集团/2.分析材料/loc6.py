import pdfplumber, os, re, sys

BASE = "/Users/zhangarmanda/PycharmProjects/Personal/投资/2_研究框架与报告/报告/环保&水务/二级/6.宁水集团/1.原始材料"
KW = ["关键审计事项", "管理层讨论与分析", "行业格局和趋势", "公司未来发展的展望",
      "主要客户", "前五名客户", "利润分配", "现金分红", "产能", "研发投入",
      "可能面对的风险", "应收账款", "信用风险", "股权激励", "承诺", "诉讼"]

files = sys.argv[1:]
for f in files:
    p = os.path.join(BASE, f)
    if not os.path.exists(p):
        print("MISS", f); continue
    hits = {k: [] for k in KW}
    with pdfplumber.open(p) as pdf:
        n = len(pdf.pages)
        for i, pg in enumerate(pdf.pages):
            t = pg.extract_text() or ""
            for k in KW:
                if k in t:
                    hits[k].append(i + 1)
    print("=" * 20, f, "pages=", n)
    for k in KW:
        v = hits[k]
        if v:
            print("  ", k, "->", v[:8], ("(+%d)" % (len(v) - 8)) if len(v) > 8 else "")
