"""取得同日美股收盘序列与行情原始快照，保留来源和时间戳。"""
from pathlib import Path
from datetime import datetime, timezone
import csv
import json
import requests
import certifi

HERE = Path(__file__).resolve().parent
S = requests.Session()
S.trust_env = False
S.headers.update({'User-Agent': 'Mozilla/5.0', 'Referer': 'https://quote.eastmoney.com/'})
END = '2026-09-16'


def get(url, params=None):
    r = S.get(url, params=params, timeout=18, verify=certifi.where())
    r.raise_for_status()
    return r


def main():
    status = {}
    for ticker in ['WRD', 'PONY', 'AUR']:
        record = {'retrieved': datetime.now().isoformat()}
        try:
            r = get('https://push2.eastmoney.com/api/qt/stock/get',
                    {'secid': f'105.{ticker}', 'fields': 'f43,f57,f58,f59,f84,f86,f116,f117,f162,f167'})
            d = r.json()
            (HERE / f'market_snapshot_{ticker}.json').write_text(json.dumps(d, ensure_ascii=False, indent=2))
            record['snapshot_url'] = r.url
            record['snapshot'] = d.get('data')
        except Exception as exc:
            record['snapshot_error'] = str(exc)
        try:
            r = get('https://push2his.eastmoney.com/api/qt/stock/kline/get',
                    {'secid': f'105.{ticker}', 'klt': 101, 'fqt': 0, 'beg': '20241001',
                     'end': END.replace('-', ''), 'lmt': 1000, 'fields1': 'f1,f2,f3,f4,f5,f6',
                     'fields2': 'f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61'})
            d = r.json()
            (HERE / f'market_daily_{ticker}_raw.json').write_text(json.dumps(d, ensure_ascii=False))
            lines = (d.get('data') or {}).get('klines') or []
            if not lines:
                raise ValueError('空行情序列')
            rows = [x.split(',') for x in lines]
            record['daily_source'] = '东方财富不复权日线'
            record['daily_url'] = r.url
        except Exception as exc:
            record['daily_primary_error'] = str(exc)
            try:
                r = get(f'https://query1.finance.yahoo.com/v8/finance/chart/{ticker}',
                        {'range': '2y', 'interval': '1d'})
                d = r.json()
                (HERE / f'market_daily_{ticker}_yahoo_raw.json').write_text(json.dumps(d))
                chart = d['chart']['result'][0]
                quote = chart['indicators']['quote'][0]
                rows = []
                for i, stamp in enumerate(chart['timestamp']):
                    day = datetime.fromtimestamp(stamp, timezone.utc).date().isoformat()
                    if day > END or quote['close'][i] is None:
                        continue
                    rows.append([day, quote['open'][i], quote['close'][i], quote['high'][i],
                                 quote['low'][i], quote['volume'][i], '', '', '', '', ''])
                record['daily_source'] = 'Yahoo chart不复权日线（东财降级）'
                record['daily_url'] = r.url
            except Exception as exc2:
                record['daily_error'] = str(exc2)
                rows = []
        if rows:
            with (HERE / f'px_daily_{ticker}.csv').open('w') as f:
                writer = csv.writer(f)
                writer.writerow(['date', 'open', 'close', 'high', 'low', 'volume', 'amount',
                                 'amplitude', 'pct_change', 'change', 'turnover'])
                writer.writerows(rows)
            record['first'] = rows[0][0]
            record['last'] = rows[-1][0]
            record['n'] = len(rows)
            record['close'] = float(rows[-1][2])
            closes = [float(row[2]) for row in rows]
            record['price_percentile'] = sum(x < closes[-1] for x in closes) / len(closes)
            record['min'] = min(closes)
            record['max'] = max(closes)
        status[ticker] = record
        (HERE / 'market_audit.json').write_text(json.dumps(status, ensure_ascii=False, indent=2))
        print(ticker, json.dumps(record, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    main()
