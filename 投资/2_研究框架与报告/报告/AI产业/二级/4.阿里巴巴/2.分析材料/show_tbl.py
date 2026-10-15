import pandas as pd, sys, re
for f in sys.argv[1:]:
    df = pd.read_csv(f, dtype=str).fillna("")
    print(f"\n########## {f} ##########")
    for _, row in df.iterrows():
        cells = [c.strip() for c in row.tolist() if c.strip() and c.strip() not in ("|","$","%",")","(")]
        if not cells: continue
        print(" | ".join(cells)[:190])
