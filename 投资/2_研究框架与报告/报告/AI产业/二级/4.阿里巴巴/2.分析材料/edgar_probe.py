import json, ssl, urllib.request, certifi
UA = {"User-Agent": "Noah Personal Investment Research (contact: noah.research@local)"}
CTX = ssl.create_default_context(cafile=certifi.where())

def get(url):
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, context=CTX, timeout=60).read()

# 1) 用 ticker 映射确认 CIK
tk = json.loads(get("https://www.sec.gov/files/company_tickers.json"))
hits = [v for v in tk.values() if v["ticker"] in ("BABA",)]
print("ticker 映射:", hits)
cik = str(hits[0]["cik_str"]).zfill(10)
print("CIK:", cik, "|", hits[0]["title"])

# 2) submissions 列近年 20-F / 6-K
sub = json.loads(get(f"https://data.sec.gov/submissions/CIK{cik}.json"))
r = sub["filings"]["recent"]
rows = list(zip(r["form"], r["filingDate"], r["reportDate"], r["accessionNumber"], r["primaryDocument"]))
print("\n== 20-F ==")
for f, fd, rd, acc, doc in rows:
    if f == "20-F":
        print(f"{fd} report={rd} acc={acc} doc={doc}")
print("\n== 最近 12 份 6-K ==")
n = 0
for f, fd, rd, acc, doc in rows:
    if f == "6-K":
        print(f"{fd} report={rd} acc={acc} doc={doc}")
        n += 1
        if n >= 12: break
print("\n表单种类统计:", {x: r["form"].count(x) for x in sorted(set(r["form"]))})
json.dump(sub, open("edgar_submissions_BABA.json", "w"))
print("\n[OK] submissions 已落盘")
