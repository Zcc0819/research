#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""第二轮定位+导出：经营主题/用户数据/风险/集中度/发展战略"""
import glob, os, re
import pdfplumber

RAW = "../1.原始材料"
OUT = "pdf_extracts"
KEYWORDS = {
    "用户数据": ["月度活跃设备", "月活跃", "累计年度付费"],
    "发展战略与经营计划": ["公司发展战略"],
    "风险": ["可能面对的风险", "可能面临的风险"],
    "客户集中度": ["前五名客户"],
}

pdfs = sorted(glob.glob(os.path.join(RAW, "*年年度报告.pdf")))
for path in pdfs:
    m = re.search(r"(20\d\d)年年度报告", os.path.basename(path))
    fy = m.group(1)
    hits = {k: [] for k in KEYWORDS}
    with pdfplumber.open(path) as pdf:
        n = len(pdf.pages)
        for i in range(min(n, 90)):  # 前90页足够覆盖MD&A
            try:
                text = pdf.pages[i].extract_text() or ""
            except Exception:
                continue
            for k, kws in KEYWORDS.items():
                if any(kw in text for kw in kws):
                    hits[k].append(i + 1)
        # 每类取前2个命中页
        pages = []
        for k, ps in hits.items():
            pages += ps[:2]
        pages = sorted(set(pages))
        print(fy, {k: v[:4] for k, v in hits.items()}, "-> dump", pages)
        buf = []
        for p in pages:
            buf.append(f"===== 第{p}页 =====")
            buf.append(pdf.pages[p - 1].extract_text() or "")
        with open(os.path.join(OUT, f"{fy}_补充页.txt"), "w") as f:
            f.write("\n".join(buf))
print("done")
