import json, os, re, ssl, time, urllib.request, certifi
UA={"User-Agent":"Noah Personal Investment Research (contact: noah.research@local)"}
CTX=ssl.create_default_context(cafile=certifi.where()); CIK="1577552"
BASE="https://www.sec.gov/Archives/edgar/data/"+CIK; RAW="../1.原始材料"
def get(u,retry=3):
    for i in range(retry):
        try: return urllib.request.urlopen(urllib.request.Request(u,headers=UA),context=CTX,timeout=120).read()
        except Exception:
            if i==retry-1: raise
            time.sleep(2*(i+1)**2)
def to_text(b):
    s=b.decode("utf-8","ignore")
    s=re.sub(r"(?is)<(script|style).*?</\1>"," ",s)
    s=re.sub(r"(?i)</t[dh]>"," | ",s); s=re.sub(r"(?i)</tr>","\n",s)
    s=re.sub(r"(?i)<br\s*/?>","\n",s); s=re.sub(r"(?i)</p>|</div>|</h[1-6]>","\n",s)
    s=re.sub(r"(?s)<[^>]+>","",s)
    for a,b2 in [("&nbsp;"," "),("&amp;","&"),("&#8217;","'"),("&#8220;",'"'),("&#8221;",'"'),("&#160;"," "),("&#8212;","—"),("&#8211;","-"),("&#39;","'")]: s=s.replace(a,b2)
    s=re.sub(r"[ \t]+"," ",s); return re.sub(r"\n\s*\n\s*\n+","\n\n",s)

sub=json.load(open("edgar_submissions_BABA.json")); r=sub["filings"]["recent"]
rows=list(zip(r["form"],r["filingDate"],r["reportDate"],r["accessionNumber"],r["primaryDocument"]))
# 季度业绩发布窗口：每年 2/5/8/11 月
MONTHS=("2025-08","2025-11","2026-02","2026-05","2026-08")
KEY=re.compile(r"(?i)Alibaba Group Announce")
targets=[]
for f,fd,rd,acc,doc in rows:
    if f!="6-K": continue
    if fd[:7] not in MONTHS and fd[:7] not in ("2026-06","2026-08"): continue
    targets.append((fd,acc))
print(f"候选 6-K {len(targets)} 份，逐份查 exhibit …")
hits=[]
for fd,acc in targets:
    accn=acc.replace("-","")
    try: idx=json.loads(get(f"{BASE}/{accn}/index.json"))
    except Exception as e: print(f" [失败]{fd}"); continue
    ex=[it["name"] for it in idx["directory"]["item"] if re.search(r"(?i)ex99|ex-99",it["name"])]
    if not ex or len(ex)>6: time.sleep(0.15); continue
    try: t=to_text(get(f"{BASE}/{accn}/{ex[0]}"))
    except Exception: time.sleep(0.15); continue
    head=re.sub(r"\s+"," ",t[:600])
    if KEY.search(head) or re.search(r"(?i)Quarter \d{4} Results|CMC List",head):
        title=re.sub(r"(?i).*?(Alibaba Group Announce[sd]?[^.]{0,90}|Inclusion of Alibaba[^.]{0,60}).*",r"\1",head)[:100]
        name=f"6-K_{fd}_{ex[0]}"
        open(f"{RAW}/c_临时公告/{name}","wb").write(get(f"{BASE}/{accn}/{ex[0]}"))
        slug=re.sub(r"[^A-Za-z0-9]+","_",title)[:48]
        open(f"html_extracts/6K_{fd}_{slug}.txt","w").write(t)
        print(f"  ★ {fd} :: {title}")
        hits.append({"date":fd,"acc":acc,"exhibit":ex[0],"title":title,"txt":f"html_extracts/6K_{fd}_{slug}.txt"})
    time.sleep(0.2)
json.dump(hits,open("edgar_6k_index.json","w"),ensure_ascii=False,indent=1)
print(f"\n命中 {len(hits)} 份实质公告")
# FY2025 文本完整性检查
for tag in ("FY2026","FY2025"):
    t=open(f"html_extracts/20F_{tag}.txt").read()
    print(f"{tag}: {len(t)/1e6:.2f}M字符  'Cloud Intelligence' 出现 {t.count('Cloud Intelligence')} 次  'Segment' {t.count('Segment')}")
