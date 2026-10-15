"""修 FY2025 20-F 正文缺失：列出该 filing 的全部文档，把主文档之外的正文分片一并下载。

背景：`20-F_FY2025_baba-20250331.htm` 仅 7.7MB 但 'Cloud Intelligence' 仅 9 次、
'free cash' 0 次，而 FY2026 同口径为 43 / 6 次——说明 FY2025 的 20-F 正文被
EDGAR 拆成多个文件，主文档只含一部分。
"""
import json
import re
import ssl
import time
import urllib.request
from pathlib import Path

import certifi

UA = {"User-Agent": "Noah Personal Investment Research (contact: noah.research@local)"}
CTX = ssl.create_default_context(cafile=certifi.where())
CIK = "1577552"
BASE = Path(__file__).resolve().parent


def get(url, retry=3):
    for i in range(retry):
        try:
            req = urllib.request.Request(url, headers=UA)
            return urllib.request.urlopen(req, context=CTX, timeout=120).read()
        except Exception:
            if i == retry - 1:
                raise
            time.sleep(2 * (i + 1) ** 2)


subs = json.loads(get(f"https://data.sec.gov/submissions/CIK{int(CIK):010d}.json"))
r = subs["filings"]["recent"]
targets = []
for form, acc, date, doc in zip(r["form"], r["accessionNumber"], r["filingDate"], r["primaryDocument"]):
    if form == "20-F" and date.startswith("2025"):
        targets.append((acc, date, doc))
print("命中 20-F：", targets)

for acc, date, doc in targets:
    accn = acc.replace("-", "")
    idx = json.loads(get(f"https://www.sec.gov/Archives/edgar/data/{CIK}/{accn}/index.json"))
    items = idx["directory"]["item"]
    print(f"\n{acc} ({date}) 共 {len(items)} 个文件：")
    for it in sorted(items, key=lambda x: -int(x.get("size") or 0)):
        size = int(it.get("size") or 0)
        if size > 200_000 and it["name"].endswith((".htm", ".html")):
            print(f"  {it['name']:<45} {size/1e6:.2f}MB")
            time.sleep(0.3)
            b = get(f"https://www.sec.gov/Archives/edgar/data/{CIK}/{accn}/{it['name']}")
            s = b.decode("utf-8", "ignore")
            n_cloud = len(re.findall(r"(?i)Cloud Intelligence", s))
            n_fcf = len(re.findall(r"(?i)free cash", s))
            print(f"      -> Cloud Intelligence={n_cloud}  free cash={n_fcf}")
            out = BASE.parent / "1.原始材料" / "a_年报" / f"20-F_FY2025_part_{it['name']}"
            out.write_bytes(b)
