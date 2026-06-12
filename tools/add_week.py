#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""주간 스냅샷 행 자동 추가 (2026 시트).
Usage: python3 add_week.py "<YSLim_Investing Summary.xlsx>" [--week W2625] [--test]
- 직전 금요일 종가(stooq.com)와 전주 대비 변동률을 마지막 W블록 2행 아래에 추가.
- 같은 주차 라벨이 이미 있으면 아무것도 하지 않음(중복 방지).
- 색상은 파일에 설정된 조건부 서식이 자동 적용(상승 빨강/하락 파랑).
"""
import sys, re, zipfile, shutil, datetime as dt, urllib.request, io, csv, html

XLSX = sys.argv[1]
ARGS = sys.argv[2:]
TEST = '--test' in ARGS
SHEET = 'xl/worksheets/sheet4.xml'   # 2026 sheet
NAMES = ['Dow','S&P500','나스닥','TSLA','SOXX','러셀2000','BMNR']
SYMB  = ['^dji','^spx','^ndq','tsla.us','soxx.us','^rut','bmnr.us']

def last_friday(today=None):
    d = today or dt.date.today()
    while d.weekday() != 4: d -= dt.timedelta(days=1)
    return d

def fetch_close(sym, on_or_before):
    url = f'https://stooq.com/q/d/l/?s={sym}&i=d'
    try:
        raw = urllib.request.urlopen(url, timeout=30).read().decode()
        rows = list(csv.reader(io.StringIO(raw)))
        best = None
        for r in rows[1:]:
            try:
                d = dt.date.fromisoformat(r[0]); c = float(r[4])
            except Exception: continue
            if d <= on_or_before: best = c
        return best
    except Exception as e:
        print(f'  fetch fail {sym}: {e}'); return None

def col_letter(idx):  # 1->A
    s=''
    while idx: idx, r = divmod(idx-1, 26); s = chr(65+r)+s
    return s

def main():
    fri = last_friday()
    iso = fri.isocalendar()
    week = f'W{str(fri.year)[2:]}{iso[1]:02d}'
    for a in ARGS:
        if a.startswith('W') and len(a)==5: week=a
    z = zipfile.ZipFile(XLSX)
    sx = z.read(SHEET).decode()
    # dedupe: label already present?
    if f'>{week}<' in sx or f'<t>{week}</t>' in sx:
        print(f'{week} already exists; nothing to do'); return 0
    # find last W-row: rows whose first cell holds inline/shared string starting with W26 - simpler: track max row number with numeric cells in D
    rownums = [int(m) for m in re.findall(r'<row r="(\d+)"', sx)]
    maxr = max(rownums); newr = maxr + 2
    # previous closes: parse last snapshot row = last row containing cells in D and G and J with numbers
    prev = {}
    lastWrow_xml = None
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sx, re.S):
        body = rm.group(2)
        cells = dict(re.findall(r'<c r="([A-Z]+)\d+"[^>]*>(?:<f>[^<]*</f>)?<v>([^<]*)</v></c>', body))
        if all(k in cells for k in ('D','G','J','M','P','S','V')):
            try:
                prev = {NAMES[i]: float(cells[c]) for i,c in enumerate(['D','G','J','M','P','S','V'])}
                lastWrow_xml = body
            except ValueError: pass
    if not prev: print('previous snapshot row not found'); return 1
    # style ids from last snapshot row
    sty = dict(re.findall(r'<c r="([A-Z]+)\d+" s="(\d+)"', lastWrow_xml or ''))
    print(f'week={week} friday={fri} newrow={newr}')
    closes = {}
    for n, s in zip(NAMES, SYMB):
        closes[n] = 123.45 if TEST else fetch_close(s, fri)
        print(f'  {n}: {closes[n]}')
    cells_xml = []
    def cell(colL, val=None, text=None):
        s = f' s="{sty[colL]}"' if colL in sty else ''
        if text is not None:
            return f'<c r="{colL}{newr}"{s} t="inlineStr"><is><t>{html.escape(text)}</t></is></c>'
        if val is None: return ''
        return f'<c r="{colL}{newr}"{s}><v>{val}</v></c>'
    cells_xml.append(cell('A', text=week))
    base = 3  # C
    for i, n in enumerate(NAMES):
        cN, cV, cP = col_letter(base+i*3), col_letter(base+i*3+1), col_letter(base+i*3+2)
        cells_xml.append(cell(cN, text=n))
        if closes[n] is not None:
            cells_xml.append(cell(cV, val=repr(closes[n])))
            if prev.get(n): cells_xml.append(cell(cP, val=repr(closes[n]/prev[n]-1)))
    newrow = f'<row r="{newr}">' + ''.join(cells_xml) + '</row>'
    sx2 = sx.replace('</sheetData>', newrow + '</sheetData>', 1)
    sx2 = re.sub(r'(<dimension ref="A1:[A-Z]+)\d+"', lambda m: f'{m.group(1)}{newr}"', sx2, 1)
    tmp = XLSX + '.tmp.xlsx'
    zo = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in z.infolist():
        zo.writestr(it, sx2 if it.filename == SHEET else z.read(it.filename))
    zo.close(); z.close()
    import openpyxl; openpyxl.load_workbook(tmp, read_only=True)  # validate
    shutil.move(tmp, XLSX)
    print('row added OK')
    return 0

sys.exit(main())
