#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""宁水集团 A 股年报关键页定位与导出（Step 3 精读取数用）。

用法：
    python extract_pages.py locate            # 全部年报按关键词定位，打印命中页
    python extract_pages.py dump 2025 12 13    # 导出指定年份指定页到 pdf_extracted/
"""
import os
import re
import sys
import warnings
import logging

warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)
import pdfplumber  # noqa: E402

RAW = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "1.原始材料")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf_extracted")

KEYWORDS = {
    "div": ["利润分配预案", "每10股派", "现金红利"],
    "seg": ["主营业务分行业情况", "主营业务分产品情况", "分行业、分产品"],
    "cust": ["前五名客户", "前五名供应商", "前五大客户"],
    "age": ["账龄", "1年以内（含1年）"],
    "nonrec": ["非经常性损益项目", "非经常性损益"],
    "kam": ["关键审计事项"],
    "rd": ["研发投入情况", "研发人员"],
    "gov": ["政府补助"],
    "ind": ["行业格局和趋势", "公司关于公司未来发展的讨论与分析", "经营情况讨论与分析"],
    "bad": ["坏账准备", "预期信用损失"],
    "pl": ["合并利润表"],
    "bs": ["合并资产负债表"],
    "cf": ["合并现金流量表"],
    "risk": ["可能面对的风险"],
}


def years():
    out = []
    for f in sorted(os.listdir(RAW)):
        m = re.match(r"(\d{4})年年度报告\.pdf$", f)
        if m:
            out.append((m.group(1), os.path.join(RAW, f)))
    return out


def locate():
    for y, path in years():
        with pdfplumber.open(path) as pdf:
            hits = {}
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                for tag, kws in KEYWORDS.items():
                    if any(k in t for k in kws):
                        hits.setdefault(tag, []).append(i + 1)
            print(f"\n### {y} 年报 共 {len(pdf.pages)} 页")
            for tag in KEYWORDS:
                ps = hits.get(tag, [])
                if ps:
                    print(f"  {tag:7s} p.{','.join(map(str, ps[:20]))}{'...' if len(ps) > 20 else ''}")
                else:
                    print(f"  {tag:7s} —")


def dump(year, pages, tag="p"):
    os.makedirs(OUT, exist_ok=True)
    path = dict(years())[year]
    with pdfplumber.open(path) as pdf:
        for p in pages:
            txt = pdf.pages[p - 1].extract_text() or "(无可提取文本)"
            fn = os.path.join(OUT, f"FY{year}_{tag}_p{p}.txt")
            with open(fn, "w") as fh:
                fh.write(txt)
            print(f"=== FY{year} p.{p} ===")
            print(txt)


if __name__ == "__main__":
    if sys.argv[1] == "locate":
        locate()
    else:
        y = sys.argv[1]
        tag = sys.argv[2]
        dump(y, [int(x) for x in sys.argv[3:]], tag)
