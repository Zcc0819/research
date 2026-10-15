"""整理限定页段、结构化交叉核验与已读取网页的降级缓存。"""
from pathlib import Path
import json
import re
import hashlib
import shutil

HERE = Path(__file__).resolve().parent
RAW = HERE.parent / '1.原始材料'
TEXT = HERE / 'source_text'
CACHE = Path('/Users/zhangarmanda/.qoder/cache/projects/Personal-9556b046/agent-tools/9e70065a')
FALLBACKS = {
    'PONY_2026Q1': ('fbd42e7d.txt', 'https://ir.pony.ai/news-releases/news-release-details/pony-ai-inc-reports-first-quarter-2026-financial-results'),
    'PONY_2026Q2': ('0d96c397.txt', 'https://ir.pony.ai/news-releases/news-release-details/pony-ai-inc-reports-second-quarter-2026-financial-results-total'),
    'BIDU_2026Q2': ('8604b198.txt', 'https://ir.baidu.com/news-releases/news-release-details/baidu-announces-second-quarter-2026-results/'),
}
manifest_path = HERE / 'source_manifest.json'
manifest = json.loads(manifest_path.read_text())
for name, (cache, url) in FALLBACKS.items():
    dest = RAW / f'{name}_webfetch.txt'
    shutil.copyfile(CACHE / cache, dest)
    manifest[dest.name] = {'url': url, 'retrieved': '2026-09-17', 'method': 'WebFetch正文缓存；requests超时降级，非原始HTML', 'sha256': hashlib.sha256(dest.read_bytes()).hexdigest()}
manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))

facts = json.loads((RAW / 'edgar_companyfacts.json').read_text())
report = {'entity': facts.get('entityName'), 'namespaces': list(facts['facts']), 'anchors': {}}
for ns, items in facts['facts'].items():
    for tag, entry in items.items():
        if tag in ['Revenue', 'RevenueFromContractsWithCustomers', 'ProfitLoss', 'ProfitLossAttributableToOwnersOfParent', 'Assets', 'Equity', 'CashFlowsFromUsedInOperatingActivities', 'NetCashFlowsFromUsedInOperatingActivities']:
            report['anchors'][ns + ':' + tag] = entry
(HERE / 'structured_anchors.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
checks = []
for tag, expected in [('Revenue', 361134000), ('ProfitLossAttributableToOwnersOfParent', -2516808000), ('Assets', 7693770000)]:
    values = facts['facts']['ifrs-full'].get(tag, {}).get('units', {}).get('CNY', [])
    found = [v for v in values if v['end'] == '2024-12-31' and v.get('form') == '20-F' and (tag == 'Assets' or v.get('start') == '2024-01-01')]
    checks.append({'tag': tag, 'pdf_CNY': expected, 'structured': found, 'equal': bool(found) and all(v['val'] == expected for v in found)})
(HERE / 'anchor_checks.json').write_text(json.dumps(checks, ensure_ascii=False, indent=2))
print('三锚点核验', [(x['tag'], x['equal']) for x in checks])
sub = json.loads((RAW / 'edgar_submissions.json').read_text())['filings']['recent']
print('20-F/424B4提交', [(sub['form'][i], sub['filingDate'][i], sub['reportDate'][i]) for i in range(len(sub['form'])) if sub['form'][i] in ['20-F', '424B4']])

wanted = {
    '2024_US_424B4': [282, 291, 310],
    '2025_HK_prospectus': [],
    'FY2024_20F': [171, 172],
    'FY2025_20F': [111, 171, 173, 174, 178, 179],
}
terms = ['unaudited condensed consolidated statements of cash flows', 'largest suppliers', 'largest customers', '2025', '2024']
for name, nums in wanted.items():
    pieces = re.split(r'===== PDF PAGE (\d+) =====', (TEXT / f'{name}.txt').read_text())
    pages = {int(pieces[i]): pieces[i+1] for i in range(1, len(pieces), 2)}
    if name == '2025_HK_prospectus':
        nums = [n for n, text in pages.items() if ('Bosch' in text and '55.3%' in text) or ('majority-controlled' in text and 'Yuji' in text) or ('378' in text and 'tax' in text.lower() and n > 550)]
    out = '\n'.join(f'===== PDF PAGE {n} =====\n{pages[n]}' for n in nums)
    (TEXT / f'{name}_supplement.txt').write_text(out)
    print(name, '定向摘页', nums)
