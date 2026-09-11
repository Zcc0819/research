import akshare as ak, time, csv
PEERS={"603700":"宁水集团","603956":"威派格","300066":"三川智慧","300259":"新天科技","300371":"汇中股份"}
rows=[]
for code,name in PEERS.items():
    for attempt in range(3):
        try:
            df=ak.stock_financial_analysis_indicator(symbol=code, start_year="2024")
            break
        except Exception as e:
            print(name,"retry",e); time.sleep(2*(attempt+1)**2); df=None
    if df is None: continue
    df=df[df["日期"].astype(str).str.endswith("12-31")]
    for _,r in df.iterrows():
        rows.append([name,code,str(r["日期"]),r.get("加权净资产收益率(%)"),r.get("销售毛利率(%)"),r.get("销售净利率(%)"),r.get("每股净资产_调整前(元)"),r.get("摊薄每股收益(元)"),r.get("应收账款周转天数(天)"),r.get("资产负债率(%)")])
    time.sleep(1.5)
hdr=["名称","代码","日期","加权ROE%","毛利率%","净利率%","BVPS","EPS","应收周转天数","资产负债率%"]
with open("peer_fin_annual.csv","w",newline="") as f:
    w=csv.writer(f); w.writerow(hdr); w.writerows(rows)
print("\t".join(hdr))
for r in rows: print("\t".join(str(x) for x in r))
