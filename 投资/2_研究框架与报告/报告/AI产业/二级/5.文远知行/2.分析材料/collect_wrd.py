"""下载公开披露并生成可定位的文本；不覆盖已缓存的原始文件。"""
from pathlib import Path
import argparse
import hashlib
import json
import re
import subprocess
import time
from datetime import datetime
from urllib.parse import urljoin

import certifi
import requests
from bs4 import BeautifulSoup

HERE = Path(__file__).resolve().parent
RAW = HERE.parent / '1.原始材料'
TEXT = HERE / 'source_text'
SESSION = requests.Session()
SESSION.trust_env = False
SESSION.headers['User-Agent'] = 'Personal investment research contact research@example.com'
SOURCES = {
    '2026H1_interim.pdf': 'https://www.hkexnews.hk/listedco/listconews/sehk/2026/0914/2026091401019.pdf',
    'policy_mot_2023.html': 'https://xxgk.mot.gov.cn/2020/jigou/ysfws/202312/t20231205_3962490.html',
    'BIDU_2026Q2.html': 'https://ir.baidu.com/news-releases/news-release-details/baidu-announces-second-quarter-2026-results/',
    'FY2025_20F.pdf': 'https://ir.weride.ai/system/files-encrypted/nasdaq_kms/assets/2026/04/23/10-08-17/20-F.pdf',
    'FY2024_20F.pdf': 'https://ir.weride.ai/static-files/72a895d1-7999-4f63-9d97-476c54e7ecab',
    '2024_US_424B4.pdf': 'https://ir.weride.ai/static-files/406cbf25-bb49-4020-b913-f9cca1283c6e',
    '2025_HK_prospectus.pdf': 'https://www1.hkexnews.hk/listedco/listconews/sehk/2025/1028/2025102800019.pdf',
    '2026Q2_release.html': 'https://ir.weride.ai/news-releases/news-release-details/accelerating-european-expansion-through-proven-asset-light-model',
    '2026Q1_release.html': 'https://ir.weride.ai/news-releases/news-release-details/wrd-30-becomes-chinas-first-four-time-urban-intelligent-driving',
    '2025Q2_release.html': 'https://ir.weride.ai/news-releases/news-release-details/weride-accelerates-global-growth-robotaxi-revenue-grew-8367',
    'quarterly_index.html': 'https://ir.weride.ai/financials-filings/quarterly-results',
    'annual_index.html': 'https://ir.weride.ai/financials-filings/annual-reports',
    'edgar_submissions.json': 'https://data.sec.gov/submissions/CIK0001867729.json',
    'edgar_companyfacts.json': 'https://data.sec.gov/api/xbrl/companyfacts/CIK0001867729.json',
}


def download(name, url):
    RAW.mkdir(parents=True, exist_ok=True)
    TEXT.mkdir(parents=True, exist_ok=True)
    path = RAW / name
    if path.exists():
        content = path.read_bytes()
    else:
        error = None
        for retry in range(3):
            try:
                response = SESSION.get(url, timeout=65, verify=certifi.where())
                response.raise_for_status()
                content = response.content
                if name.endswith('.pdf') and not content.startswith(b'%PDF'):
                    raise ValueError('响应不是PDF')
                if name.endswith('.json'):
                    json.loads(content)
                path.write_bytes(content)
                break
            except Exception as exc:
                error = str(exc)
                time.sleep(2 * (retry + 1))
        else:
            print('FAILED', name, error, flush=True)
            return {'name': name, 'url': url, 'error': error}
    result = {'name': name, 'url': url, 'retrieved': datetime.now().isoformat(),
              'bytes': len(content), 'sha256': hashlib.sha256(content).hexdigest()}
    target = TEXT / (Path(name).stem + '.txt')
    if name.endswith('.pdf') and not target.exists():
        process = subprocess.run(['pdftotext', '-layout', str(path), '-'],
                                 capture_output=True, check=True)
        pages = process.stdout.decode('utf-8').split('\f')
        if not pages[-1].strip():
            pages.pop()
        target.write_text('\n'.join(f'\n===== PDF PAGE {i} =====\n{text}' for i, text in enumerate(pages, 1)))
        hits = {}
        for term in ['CONSOLIDATED STATEMENTS', 'CONSOLIDATED BALANCE', 'cash flows', 'Cash Flows',
                     'critical audit', 'Critical Audit', 'material weakness', 'going concern',
                     'Robotaxi', 'research and development', 'major customers', 'share-based',
                     'IFRS', 'unit economics', 'use of proceeds', 'beneficial ownership']:
            hits[term] = [i for i, text in enumerate(pages, 1) if term.lower() in text.lower()]
        (TEXT / (Path(name).stem + '_pages.json')).write_text(json.dumps(hits, indent=2))
        result['pages'] = len(pages)
    elif name.endswith('.html') and not target.exists():
        soup = BeautifulSoup(content, 'html.parser')
        for node in soup(['script', 'style']):
            node.decompose()
        target.write_text(soup.get_text('\n', strip=True))
        if 'index' in name:
            links = [{'text': a.get_text(' ', strip=True), 'url': urljoin(url, a['href'])}
                     for a in soup.select('a[href]')
                     if any(t in a.get('href', '') for t in ['static-files', 'news-release', '/node/', '.pdf'])]
            (TEXT / (Path(name).stem + '_links.json')).write_text(json.dumps(links, ensure_ascii=False, indent=2))
    print('OK', name, result.get('pages', ''), len(content), flush=True)
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--name')
    parser.add_argument('--url')
    args = parser.parse_args()
    manifest_path = HERE / 'source_manifest.json'
    manifest = json.loads(manifest_path.read_text()) if manifest_path.exists() else {}
    sources = {args.name: args.url} if args.name else SOURCES
    for name, url in sources.items():
        manifest[name] = download(name, url)
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
