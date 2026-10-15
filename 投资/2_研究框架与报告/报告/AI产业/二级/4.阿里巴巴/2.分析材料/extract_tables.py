import re, os, sys, io
import pandas as pd
SRC = "../1.原始材料/a_年报"
OUT = "tables"; os.makedirs(OUT, exist_ok=True)
KEYS = [("segment_revenue", re.compile(r"(?is)Cloud Intelligence Group")),]

def clean(df):
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df = df.loc[:, ~(df.astype(str).apply(lambda c: c.str.fullmatch(r"\s*(\||\$|%|nan|\)|\()?\s*", na=True)).all())]
    return df

for fn in sorted(os.listdir(SRC)):
    if not fn.endswith(".htm"): continue
    tag = fn.split("_")[1]
    html = open(os.path.join(SRC, fn), encoding="utf-8", errors="ignore").read()
    tables = re.findall(r"(?is)<table[^>]*>.*?</table>", html)
    print(f"\n=== {tag}: {len(tables)} 张表 ===")
    hit = 0
    for i, t in enumerate(tables):
        txt = re.sub(r"(?s)<[^>]+>", " ", t)
        if "Cloud Intelligence Group" not in txt: continue
        # 只要含分部名 + (Revenue 或 EBITA) 的表
        if not re.search(r"(?i)revenue|EBITA", txt): continue
        try:
            dfs = pd.read_html(io.StringIO(t))
        except Exception as e:
            continue
        for df in dfs:
            df = clean(df)
            if df.shape[0] < 3 or df.shape[1] < 3: continue
            kind = "EBITA" if re.search(r"(?i)EBITA", txt) else "REV"
            path = f"{OUT}/{tag}_t{i:03d}_{kind}.csv"
            df.to_csv(path, index=False)
            hit += 1
            first_col = " / ".join(str(x)[:22] for x in df.iloc[:, 0].head(9).tolist())
            print(f"  [{kind}] t{i:03d} {df.shape} -> {os.path.basename(path)}")
            print(f"      行头: {first_col}")
    print(f"  命中 {hit} 张")
