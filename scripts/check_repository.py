#!/usr/bin/env python3
"""Check repository links, original data, exports and constructed metrics.

Uses disposable outputs. Does not recalculate formulas or visit research sites.
"""
import argparse
import collections
import csv
import datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from urllib.parse import unquote, urlsplit
from refresh_productmatch import XLSX, colname

ROOT = Path(__file__).resolve().parents[1]


def rows(path):
    with path.open(encoding='utf-8-sig', newline='') as f:
        return list(csv.reader(f))


def run(*args):
    result = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT,
                            text=True, capture_output=True, check=True)
    return result.stdout


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', type=Path, default=ROOT / '.artifacts/repository-checks.json')
    args = parser.parse_args()
    tests = []
    workbook = ROOT / 'ProductMatch.xlsx'
    initial_hash = hashlib.sha256(workbook.read_bytes()).hexdigest()
    x = XLSX(workbook)
    for name, filename, count in [('MASTER RAW','master_raw.csv',200), ('SUPPLIER RAW','supplier_raw.csv',250)]:
        raw = rows(ROOT / 'data/raw' / filename)
        assert len(raw) - 1 == count
        cells = x.read(name)
        for r, row in enumerate(raw[1:], 7):
            for c, value in enumerate(row, 1):
                assert str(cells.get(f'{colname(c)}{r}', '')) == value, (name,r,c)
    assert rows(ROOT / 'data/raw/master_raw.csv')[1][2] == '0010001'
    tests.append('450 raw rows agree cell-for-cell, including leading-zero identifiers and whitespace')

    with tempfile.TemporaryDirectory(prefix='productmatch-repository-') as temp:
        tmp = Path(temp)
        run('scripts/refresh_productmatch.py','--workbook',workbook,'--export-cached',tmp/'exports')
        for delivered in (ROOT / 'exports').glob('*.csv'):
            assert rows(delivered) == rows(tmp / 'exports' / delivered.name), delivered.name
        tests.append('Documented export CLI reproduces all four delivered CSV datasets')
        run('scripts/generate_demo_data.py','--output-dir',tmp)
        for relative in ['data/raw/master_raw.csv','data/raw/supplier_raw.csv',
                         'data/reference/research_evidence.csv','data/evaluation/ground_truth.csv']:
            assert rows(ROOT / relative) == rows(tmp / relative), relative
        generated = json.loads((tmp/'data/generated/demo_data.json').read_text(encoding='utf-8'))
        delivered = json.loads((ROOT/'data/generated/demo_data.json').read_text(encoding='utf-8'))
        assert generated == delivered
        run('scripts/refresh_productmatch.py','--data',tmp/'data/generated/demo_data.json',
            '--output',tmp/'data/generated/candidate_snapshot.json')
        candidate = json.loads((tmp/'data/generated/candidate_snapshot.json').read_text(encoding='utf-8'))
        baseline = json.loads((ROOT/'data/generated/candidate_snapshot.json').read_text(encoding='utf-8'))
        assert candidate['masterSnapshot'] == baseline['masterSnapshot']
        assert [r[:8] for r in candidate['rows']] == [r[:8] for r in baseline['rows']]
        tests.append('Generation reproduces raw data, research, labels, builder inputs and candidate scores; timestamps excluded')

    truth_rows = rows(ROOT/'data/evaluation/ground_truth.csv')[1:]
    labels = {r[0]:r[1] for r in truth_rows if not r[3]}
    assert len(truth_rows) == 250 and len(labels) == 235
    positives = sum(target != 'NONE' for target in labels.values())
    with (ROOT/'exports/matching_review.csv').open(encoding='utf-8-sig',newline='') as f:
        review = list(csv.DictReader(f))
    eligible = [r for r in review if r['Supplier row ID'] in labels]
    suggestions = [r for r in eligible if r['Suggested master ID']]
    correct = sum(r['Suggested master ID'] == labels[r['Supplier row ID']] for r in suggestions)
    accepted = [r for r in review if r['Review outcome'] == 'Accepted']
    assert (len(suggestions),correct,positives) == (217,204,215)
    assert len(accepted) == 124
    assert all(r['Final product ID'] == labels[r['Supplier row ID']] for r in accepted)
    outcomes = dict(collections.Counter(r['Review outcome'] for r in review))
    assert outcomes == {'Accepted':124,'Needs review':88,'Rejected':5,'Excluded duplicate':15,'Unmatched':18}
    final = rows(ROOT/'exports/final_clean_catalog.csv')
    exceptions = rows(ROOT/'exports/exceptions.csv')
    assert len(final)-1 == 124 and len(exceptions)-1 == 111
    assert len({r[0] for r in final[1:]}) == 124 and final[1][3] == '0010001'
    required = [3,4,5,6,9,10,11,12,13,14]
    populated = sum(str(r[c]).strip() != '' for r in rows(ROOT/'data/raw/supplier_raw.csv')[1:] for c in required)
    assert populated == 2463
    tests.append('Independent constructed-label evaluation and 124 + 111 + 15 supplier partition agree')
    research = rows(ROOT/'data/reference/research_evidence.csv')[1:]
    assert len(research) == 22
    assert len({r[1] for r in research if r[3]=='Specification' and r[6]=='Verified'}) == 15
    assert {r[11] for r in research if r[11]} == {'2026-11-11'}
    tests.append('Research register retains 15 distinct verified real models and original access dates')

    checked_links = 0
    markdown = [ROOT/'README.md',ROOT/'data/README.md',*(ROOT/'docs').rglob('*.md')]
    for path in markdown:
        text = path.read_text(encoding='utf-8')
        assert not re.search(r'sandbox:|chatgpt\.com/(?:backend|c/)|[A-Za-z]:\\|/workspace/scratch/',text),path
        for target in re.findall(r'!?\[[^\]]*\]\(([^\s)]+)(?:\s+[^)]*)?\)',text):
            parsed = urlsplit(target.strip('<>'))
            if parsed.scheme or not parsed.path:
                continue
            destination = (path.parent/unquote(parsed.path)).resolve()
            assert destination.is_relative_to(ROOT) and destination.exists(),(path,target)
            checked_links += 1
    tests.append(f'{checked_links} relative documentation/image links resolve; public docs have no local/private download paths')
    assert hashlib.sha256(workbook.read_bytes()).hexdigest() == initial_hash
    evidence = {'verification_date':datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))).date().isoformat(),
                'python':sys.version.split()[0],'workbook_sha256':initial_hash,'tests':tests,
                'constructed_evaluation':{'suggestions':len(suggestions),'correct':correct,'positives':positives,'accepted':len(accepted)},
                'limitations':['No formula recalculation','No native Excel or Google Sheets UI test','Research sites not revisited']}
    args.report.parent.mkdir(parents=True,exist_ok=True)
    args.report.write_text(json.dumps(evidence,indent=2),encoding='utf-8')
    print(json.dumps({'checks_passed':len(tests),'relative_links':checked_links,'report':str(args.report)}))


if __name__ == '__main__':
    main()
