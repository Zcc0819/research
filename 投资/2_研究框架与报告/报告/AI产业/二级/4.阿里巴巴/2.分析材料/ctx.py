"""按关键词打印跨年上下文（BSD grep 的 .\\{0,N\\} 上限 255，用 Python 更稳）。

用法：python ctx.py "关键词" [窗口字符数] [文件1 文件2 ...]
"""
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent / "html_extracts"
DEFAULT = [
    "20F_FY2022.txt",
    "20F_FY2023.txt",
    "20F_FY2024.txt",
    "20F_FY2025_正文Ex15.5.txt",
    "20F_FY2026.txt",
]


def main() -> None:
    kw = sys.argv[1]
    win = int(sys.argv[2]) if len(sys.argv) > 2 else 300
    files = sys.argv[3:] or DEFAULT
    for name in files:
        p = BASE / name
        if not p.exists():
            print(f"[skip] {name}")
            continue
        s = p.read_text(errors="ignore")
        hits = list(re.finditer(kw, s, re.I))
        print(f"\n########## {name}  命中 {len(hits)} 次")
        for m in hits[:4]:
            seg = s[max(0, m.start() - win): m.end() + win]
            seg = re.sub(r"\s+", " ", seg).strip()
            print(f"  ...{seg}...\n")


if __name__ == "__main__":
    main()
