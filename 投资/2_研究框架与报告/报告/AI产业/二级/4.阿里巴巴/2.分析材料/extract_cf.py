"""从 20-F HTML 里按关键词提取表格（现金流/折旧/股本），落 CSV。

复用 extract_tables.py 的方法：iXBRL 表格不能用剥标签的方式读，
必须切出 <table> 块交给 pandas.read_html。
"""
import io
import re
import sys
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
RAW = BASE.parent / "1.原始材料" / "a_年报"
OUT = BASE / "tables"
OUT.mkdir(exist_ok=True)

# (标签, 必含关键词, 可选关键词之一)
PATTERNS = [
    ("FCF", ["Free cash flow"], ["operating activities"]),
    ("DEP", ["Depreciation"], ["amortization", "Amortization"]),
    ("SHARES", ["ordinary shares"], ["issued and outstanding", "Weighted average"]),
]


def main(fy: str) -> None:
    src = RAW / f"20F_{fy}.htm"
    if not src.exists():
        cands = sorted(RAW.glob(f"*{fy}*"))
        if not cands:
            print(f"[skip] {fy} 未找到原文")
            return
        src = cands[0]
    html = src.read_text(errors="ignore")
    tables = re.findall(r"(?is)<table[^>]*>.*?</table>", html)
    print(f"{fy}: {src.name}, {len(tables)} 张表")
    hit = 0
    for i, t in enumerate(tables):
        txt = re.sub(r"(?s)<[^>]+>", " ", t)
        for tag, musts, anys in PATTERNS:
            if not all(m in txt for m in musts):
                continue
            if anys and not any(a in txt for a in anys):
                continue
            try:
                dfs = pd.read_html(io.StringIO(t))
            except ValueError:
                continue
            for j, df in enumerate(dfs):
                if df.shape[0] < 3:
                    continue
                p = OUT / f"{fy}_t{i:03d}_{tag}_{j}.csv"
                df.to_csv(p, index=False)
                print(f"  -> {p.name} {df.shape}")
                hit += 1
    print(f"{fy}: 命中 {hit} 张")


if __name__ == "__main__":
    for fy in sys.argv[1:] or ["FY2026"]:
        main(fy)
