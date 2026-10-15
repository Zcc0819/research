import json, os, re, ssl, time, urllib.request, certifi
UA = {"User-Agent": "Noah Personal Investment Research (contact: noah.research@local)"}
CTX = ssl.create_default_context(cafile=certifi.where())
CIK = "1577552"
BASE = "https://www.sec.gov/Archives/edgar/data/" + CIK
RAW = "../1.原始材料"
os.makedirs(RAW + "/a_年报", exist_ok=True)
os.makedirs(RAW + "/c_临时公告", exist_ok=True)
os.makedirs("html_extracts", exist_ok=True)

def get(url, retry=3):
    for i in range(retry):
        try:
            req = urllib.request.Request(url, headers=UA)
            return urllib.request.urlopen(req, context=CTX, timeout=120).read()
        except Exception as e:
            if i == retry - 1: raise
            time.sleep(2 * (i + 1) ** 2)

def to_text(html_bytes):
    s = html_bytes.decode("utf-8", "ignore")
    s = re.sub(r"(?is)<(script|style).*?</\1>", " ", s)
    s = re.sub(r"(?i)</t[dh]>", " | ", s)
    s = re.sub(r"(?i)</tr>", "\n", s)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</p>|</div>|</h[1-6]>", "\n", s)
    s = re.sub(r"(?s)<[^>]+>", "", s)
    s = s.replace("&nbsp;", " ").replace("&amp;", "&").replace("&#8217;", "'").replace("&#8220;", '"').replace("&#8221;", '"').replace("&#160;", " ").replace("&#8212;", "—").replace("&#8211;", "-")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n\s*\n\s*\n+", "\n\n", s)
    return s

sub = json.load(open("edgar_submissions_BABA.json"))
r = sub["filings"]["recent"]
rows = list(zip(r["form"], r["filingDate"], r["reportDate"], r["accessionNumber"], r["primaryDocument"]))

# ---- A. 20-F 五年正文 ----
print("=== 下载 20-F ===")
want = {"2026-03-31": "FY2026", "2025-03-31": "FY2025", "2024-03-31": "FY2024", "2023-03-31": "FY2023", "2022-03-31": "FY2022"}
for f, fd, rd, acc, doc in rows:
    if f != "20-F" or rd not in want: continue
    tag = want[rd]
    accn = acc.replace("-", "")
    url = f"{BASE}/{accn}/{doc}"
    dst = f"{RAW}/a_年报/20-F_{tag}_{doc}"
    if os.path.exists(dst):
        print(f"  [skip] {tag}"); continue
    b = get(url); open(dst, "wb").write(b)
    txt = to_text(b); open(f"html_extracts/20F_{tag}.txt", "w").write(txt)
    print(f"  [OK] {tag} html={len(b)/1e6:.1f}MB text={len(txt)/1e6:.1f}M字符 filed={fd}")
    time.sleep(0.3)

# ---- B. 最近 6-K 找季度业绩公告 ----
print("\n=== 探测最近 6-K 的 exhibits（找季度业绩） ===")
found = []
cnt = 0
for f, fd, rd, acc, doc in rows:
    if f != "6-K": continue
    cnt += 1
    if cnt > 30: break
    accn = acc.replace("-", "")
    try:
        idx = json.loads(get(f"{BASE}/{accn}/index.json"))
    except Exception as e:
        print(f"  [失败] {fd} {e}"); continue
    names = [it["name"] for it in idx["directory"]["item"]]
    ex99 = [n for n in names if re.search(r"(?i)ex99|ex-99", n)]
    label = ""
    for n in ex99[:1]:
        try:
            head = to_text(get(f"{BASE}/{accn}/{n}"))[:400].replace("\n", " ")
            label = re.sub(r"\s+", " ", head)[:150]
        except Exception: pass
        time.sleep(0.2)
    hit = bool(re.search(r"(?i)quarter|results|revenue", label))
    print(f"  {fd} ex99={len(ex99)} {'★' if hit else ' '} {label[:110]}")
    if hit: found.append((fd, acc, ex99, label))
    time.sleep(0.2)
json.dump(found, open("edgar_6k_results_candidates.json", "w"), ensure_ascii=False, indent=1)
print(f"\n候选业绩公告 {len(found)} 份，已落盘 edgar_6k_results_candidates.json")
