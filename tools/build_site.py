#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""YSLim Investing Summary - static site builder (v2, single workbook).
Usage: python3 build_site.py "<YSLim_Investing Summary.xlsx>" <out_dir>
- 시트: Orientation / Summary_~2024 / 2025 / 2026 / Factset / Yield
- 흰색 폰트 셀(비공개 메모)은 사이트에서 제외된다.
"""
import sys, os, re, zipfile, datetime, json, shutil, html as H

XLSX, SITE = sys.argv[1], sys.argv[2]
TOOLS = os.path.dirname(os.path.abspath(__file__))
IMGDIR = os.path.join(SITE, 'assets', 'img')
os.makedirs(IMGDIR, exist_ok=True)
shutil.copy(os.path.join(TOOLS, 'chart.umd.js'), os.path.join(SITE, 'assets', 'chart.umd.js'))

LINKS6 = [
 ("Financial condition & BEI & Sales/Inventory", "https://fredaccount.stlouisfed.org/public/dashboard/111622", "FRED 대시보드 111622"),
 ("Yield & Bond", "https://fredaccount.stlouisfed.org/public/dashboard/111624", "FRED 대시보드 111624"),
 ("FED Balance Sheet", "https://fredaccount.stlouisfed.org/public/dashboard/111626", "FRED 대시보드 111626"),
 ("환율 (FX)", "https://fredaccount.stlouisfed.org/public/dashboard/134732", "FRED 대시보드 134732"),
 ("GDPNow", "https://www.atlantafed.org/cqer/research/gdpnow#Tab3", "Atlanta Fed"),
 ("CME FedWatch", "https://www.cmegroup.com/markets/interest-rates/cme-fedwatch", "CME Group"),
]

CSS = """
:root{--bg:#0f1419;--card:#1a2129;--border:#2a3441;--text:#e6edf3;--muted:#8b98a5;--accent:#4da3ff;}
*{box-sizing:border-box;}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI','Malgun Gothic',system-ui,sans-serif;margin:0;padding:0 16px 60px;}
header{display:flex;flex-wrap:wrap;align-items:baseline;gap:18px;padding:18px 4px;border-bottom:1px solid var(--border);margin-bottom:20px;}
header h1{font-size:19px;margin:0;}
nav{display:flex;flex-wrap:wrap;gap:4px;}
nav a{color:var(--muted);text-decoration:none;font-size:14px;padding:4px 10px;border-radius:6px;}
nav a:hover{color:var(--text);background:var(--card);}
nav a.on{color:var(--accent);font-weight:600;}
.wrap{max-width:1060px;margin:0 auto;}
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px 18px;margin-bottom:16px;}
.card h2{font-size:16px;color:var(--accent);margin:0 0 10px;}
.week{border-left:3px solid var(--accent);}
table.snap{border-collapse:collapse;width:100%;margin:6px 0 12px;font-size:13px;}
table.snap th,table.snap td{border:1px solid var(--border);padding:5px 9px;text-align:right;}
table.snap th{color:var(--muted);font-weight:600;}
.up{color:#f85149;}.dn{color:#4da3ff;}
.line{font-size:13.5px;line-height:1.65;color:var(--text);margin:2px 0;white-space:pre-wrap;}
.line a{color:var(--accent);}
img.emb{max-width:100%;border-radius:8px;border:1px solid var(--border);margin:8px 6px 8px 0;vertical-align:top;}
.imgrow{display:flex;flex-wrap:wrap;gap:8px;}
.linkgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px;}
.linkgrid a{display:block;background:var(--card);border:1px solid var(--border);border-radius:10px;padding:14px 16px;text-decoration:none;}
.linkgrid a:hover{border-color:var(--accent);}
.linkgrid .t{color:var(--text);font-weight:600;font-size:14.5px;}
.linkgrid .s{color:var(--muted);font-size:12.5px;margin-top:4px;}
.muted{color:var(--muted);font-size:12.5px;}
details{margin-bottom:14px;}
details summary{cursor:pointer;font-size:15px;color:var(--accent);padding:10px 14px;background:var(--card);border:1px solid var(--border);border-radius:10px;}
details[open] summary{border-radius:10px 10px 0 0;}
details .body{border:1px solid var(--border);border-top:none;border-radius:0 0 10px 10px;padding:12px 16px;background:#141a21;}
"""
# 상승=빨강(.up), 하락=파랑(.dn) — 한국식 컬러

def page(title, active, body):
    nav = ''.join(f'<a href="{h}" class="{"on" if k==active else ""}">{t}</a>' for t,h,k in [
        ('최신 브리핑','index.html','index'),('퀵 링크','links.html','links'),
        ('Orientation','orientation.html','orient'),('2026','archive-2026.html','2026'),
        ('2025','archive-2025.html','2025'),('2023–2024','archive-2024.html','2024'),
        ('Factset','factset.html','factset'),('Yield','yield.html','yield')])
    return f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{H.escape(title)} · YSLim Investing</title><style>{CSS}</style></head>
<body><div class="wrap"><header><h1>YSLim Investing Summary</h1><nav>{nav}</nav></header>
{body}
<p class="muted">생성: {datetime.date.today().isoformat()} · 본 자료는 개인 기록·공유용이며 투자 권유가 아닙니다. 투자 판단과 책임은 각자에게 있습니다.</p>
</div></body></html>"""

def esc(x): return H.escape(str(x))

def sheet_xml_map(zf):
    wbx = zf.read('xl/workbook.xml').decode()
    sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx)
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', zf.read('xl/_rels/workbook.xml.rels').decode()))
    return {name: rels[rid].split('/')[-1] for name, rid in sheets}

def sheet_assets(zf, sheetfile, prefix):
    imgs, links = [], {}
    relp = f'xl/worksheets/_rels/{sheetfile}.rels'
    if relp not in zf.namelist(): return imgs, links
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', zf.read(relp).decode()))
    sx = zf.read(f'xl/worksheets/{sheetfile}').decode()
    for ref, rid in re.findall(r'<hyperlink ref="([A-Z]+\d+)" r:id="(rId\d+)"', sx):
        if rid in rels: links[int(re.search(r'\d+', ref).group())] = rels[rid]
    drawing = [v for v in rels.values() if 'drawing' in v]
    if drawing:
        dname = drawing[0].split('/')[-1]
        dr = zf.read(f'xl/drawings/{dname}').decode()
        drels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', zf.read(f'xl/drawings/_rels/{dname}.rels').decode()))
        for a in re.findall(r'<xdr:(?:two|one)CellAnchor.*?</xdr:(?:two|one)CellAnchor>', dr, re.S):
            rowm = re.search(r'<xdr:from>.*?<xdr:row>(\d+)</xdr:row>', a, re.S)
            em = re.search(r'r:embed="(rId\d+)"', a)
            if not (rowm and em and em.group(1) in drels): continue
            mediafile = drels[em.group(1)].split('/')[-1]
            out = f'{prefix}_{mediafile}'
            outp = os.path.join(IMGDIR, out)
            if not os.path.exists(outp):
                with open(outp, 'wb') as f: f.write(zf.read(f'xl/media/{mediafile}'))
            imgs.append((int(rowm.group(1)) + 1, 0, f'assets/img/{out}'))
    return imgs, links

def is_white(cell):
    col = cell.font.color
    if col is None: return False
    t = getattr(col, 'type', None)
    if t == 'rgb' and str(col.rgb).upper() in ('FFFFFFFF', '00FFFFFF'): return True
    if t == 'theme' and col.theme == 0 and not getattr(col, 'tint', 0): return True
    return False

ASSETS = ['Dow','S&P500','나스닥','TSLA','SOXX','러셀2000','BMNR']
def fmt_pct(v):
    try: v = float(v)
    except: return ''
    cls = 'up' if v >= 0 else 'dn'
    return f'<span class="{cls}">{v*100:+.2f}%</span>'

def render_rows(rows, imgs, links, r0, r1, week_mode=True):
    out = []
    imap = {}
    for r, c, p in imgs:
        if r0 <= r <= r1: imap.setdefault(r, []).append(p)
    i = r0
    while i <= r1:
        row = rows[i-1] if i-1 < len(rows) else ()
        cells = [c for c in row if c is not None and str(c).strip() != '']
        first = str(row[0]).strip() if row and row[0] is not None else ''
        if week_mode and re.match(r'^W\d{4}', first):
            head, vals = [], []
            for j, name in enumerate(ASSETS):
                base = 2 + j*3
                if base+1 < len(row) and isinstance(row[base+1], (int, float)):
                    head.append(f'<th>{esc(row[base] or name)}</th>')
                    pct = fmt_pct(row[base+2]) if base+2 < len(row) else ''
                    v = row[base+1]
                    vs = f'{v:,.2f}' if isinstance(v,(int,float)) else esc(v)
                    vals.append(f'<td>{vs}<br>{pct}</td>')
            out.append(f'<h2>{esc(first)}</h2><table class="snap"><tr>{"".join(head)}</tr><tr>{"".join(vals)}</tr></table>')
        elif cells:
            txt = '  '.join(esc(c) for c in cells if not isinstance(c, datetime.datetime))
            if i in links:
                u = esc(links[i]); txt += f' <a href="{u}" target="_blank">[link]</a>'
            if txt.strip(): out.append(f'<div class="line">{txt}</div>')
        if i in imap:
            out.append('<div class="imgrow">' + ''.join(f'<a href="{p}" target="_blank"><img class="emb" src="{p}" loading="lazy" style="max-width:480px"></a>' for p in imap[i]) + '</div>')
        i += 1
    return '\n'.join(out)

def render_summary_sheet(rows, imgs, links):
    starts = []
    for i, row in enumerate(rows, 1):
        first = str(row[0]).strip() if row and row[0] is not None else ''
        if re.match(r'^W\d{4}', first): starts.append(i)
    if not starts: return [('전체', render_rows(rows, imgs, links, 1, len(rows), week_mode=False))]
    blocks, seen = [], {}
    for k, s in enumerate(starts):
        e = (starts[k+1]-1) if k+1 < len(starts) else len(rows)
        lbl = str(rows[s-1][0]).strip()
        if lbl in seen: seen[lbl] += 1; lbl = f'{lbl}({seen[lbl]})'
        else: seen[lbl] = 1
        blocks.append((lbl, render_rows(rows, imgs, links, s, e)))
    return blocks

def archive_page(blocks, latest_open=False):
    parts = []
    for k, (lbl, bh) in enumerate(reversed(blocks)):
        op = ' open' if (latest_open and k == 0) else ''
        parts.append(f'<details{op}><summary>{esc(lbl)}</summary><div class="body week">{bh}</div></details>')
    return f'<p class="muted">{len(blocks)}개 주차 · 최신순 · 제목을 누르면 펼쳐집니다</p>' + '\n'.join(parts)

# ---------- load ----------
import openpyxl
print('loading workbook (styles for privacy filter)...')
wb = openpyxl.load_workbook(XLSX, data_only=True)
zf = zipfile.ZipFile(XLSX)
smap = sheet_xml_map(zf)

def masked_rows(sheet):
    ws = wb[sheet]; out = []
    for row in ws.iter_rows():
        out.append(tuple(None if (c.value is not None and is_white(c)) else c.value for c in row))
    return out

PAGES = []
# 2026
rows26 = masked_rows('2026'); img26, lnk26 = sheet_assets(zf, smap['2026'], 'y26')
b26 = render_summary_sheet(rows26, img26, lnk26)
open(os.path.join(SITE,'archive-2026.html'),'w',encoding='utf-8').write(page('2026 Archive','2026', archive_page(b26)))
print('2026:', len(b26))
# 2025 / ~2024
rows25 = masked_rows('2025'); img25, lnk25 = sheet_assets(zf, smap['2025'], 'y25')
b25 = render_summary_sheet(rows25, img25, lnk25)
open(os.path.join(SITE,'archive-2025.html'),'w',encoding='utf-8').write(page('2025 Archive','2025', archive_page(b25)))
print('2025:', len(b25))
rows24 = masked_rows('Summary_~2024'); img24, lnk24 = sheet_assets(zf, smap['Summary_~2024'], 'y24')
b24 = render_summary_sheet(rows24, img24, lnk24)
open(os.path.join(SITE,'archive-2024.html'),'w',encoding='utf-8').write(page('2023–2024 Archive','2024', archive_page(b24)))
print('~2024:', len(b24))
# Orientation
rowsor = masked_rows('Orientation'); imgor, lnkor = sheet_assets(zf, smap['Orientation'], 'ori')
orient_html = '<div class="card">' + render_rows(rowsor, imgor, lnkor, 1, len(rowsor), week_mode=False) + '</div>'
open(os.path.join(SITE,'orientation.html'),'w',encoding='utf-8').write(page('Orientation','orient',
  '<p class="muted">지표 체계 설명서 — Orientation 시트</p>' + orient_html))
# links
lg = ''.join(f'<a href="{u}" target="_blank"><div class="t">{i+1}. {esc(t)}</div><div class="s">{esc(s)}</div></a>' for i,(t,u,s) in enumerate(LINKS6))
open(os.path.join(SITE,'links.html'),'w',encoding='utf-8').write(page('퀵 링크','links',
  f'<p class="muted">Summary 시트 상단 고정 링크 6개</p><div class="linkgrid">{lg}</div>'))
# factset / yield
fs_rows = [tuple(c.value for c in r) for r in wb['Factset'].iter_rows()]
fs = [[r[0].strftime('%Y-%m-%d')] + [float(x) if isinstance(x,(int,float)) else None for x in r[1:8]] for r in fs_rows if r and hasattr(r[0],'strftime')]
yd_rows = [tuple(c.value for c in r) for r in wb['Yield'].iter_rows()]
yh = [str(x) for x in yd_rows[0][:19]]
yd = []
for r in yd_rows[1:]:
    if r[0] is None or not hasattr(r[0],'strftime'): continue
    yd.append([r[0].strftime('%Y-%m-%d')] + [float(x) if isinstance(x,(int,float)) else None for x in r[1:19]])
CHART_HDR = '<script src="assets/chart.umd.js"></script>'
JSBASE = "Chart.defaults.color='#8b98a5';Chart.defaults.borderColor='#2a3441';"
fs_table_rows = ''.join('<tr>' + ''.join(f'<td>{("" if v is None else (f"{v:.1f}" if isinstance(v,float) and i!=4 else (f"{v*100:.1f}%" if i==4 and v is not None else v)))}</td>' for i,v in enumerate(r)) + '</tr>' for r in reversed(fs[-30:]))
fs_body = f"""{CHART_HDR}
<div class="card"><h2>Earnings Growth & Positive Guidance %</h2><div style="height:330px;position:relative"><canvas id="c1"></canvas></div></div>
<div class="card"><h2>FWD 12M PER vs 5Y/10Y 평균</h2><div style="height:330px;position:relative"><canvas id="c2"></canvas></div></div>
<div class="card"><h2>최근 30개 리포트</h2><div style="overflow-x:auto"><table class="snap"><tr><th>Date</th><th>Earn growth%</th><th>Neg G</th><th>Pos G</th><th>Pos %</th><th>FWD PER</th><th>5Y avg</th><th>10Y avg</th></tr>{fs_table_rows}</table></div></div>
<script>const F={json.dumps(fs)};{JSBASE}
const L=F.map(r=>r[0]);
new Chart(c1,{{data:{{labels:L,datasets:[{{type:'bar',label:'Earnings growth %',data:F.map(r=>r[1]),backgroundColor:F.map(r=>r[1]>=0?'rgba(248,81,73,.55)':'rgba(77,163,255,.55)')}},{{type:'line',label:'Positive guidance %(우)',data:F.map(r=>r[4]==null?null:r[4]*100),borderColor:'#d29922',pointRadius:0,yAxisID:'y2'}}]}},options:{{maintainAspectRatio:false,scales:{{x:{{ticks:{{maxTicksLimit:10}}}},y2:{{position:'right',min:0,max:100,grid:{{drawOnChartArea:false}}}}}}}}}});
new Chart(c2,{{type:'line',data:{{labels:L,datasets:[{{label:'FWD 12M PER',data:F.map(r=>r[5]),borderColor:'#4da3ff',pointRadius:0}},{{label:'5Y avg',data:F.map(r=>r[6]),borderColor:'#3fb950',pointRadius:0,borderDash:[5,4]}},{{label:'10Y avg',data:F.map(r=>r[7]),borderColor:'#d29922',pointRadius:0,borderDash:[2,3]}}]}},options:{{maintainAspectRatio:false,scales:{{x:{{ticks:{{maxTicksLimit:10}}}}}}}}}});
</script>"""
open(os.path.join(SITE,'factset.html'),'w',encoding='utf-8').write(page('Factset','factset', fs_body))
yd_last = yd[-1]
# Yield_Chart 시트의 두 차트를 Yield 데이터로 재현 (Today/월평균, Today~Today-4)
full = [r for r in yd if all(v is not None for v in r[1:14])]
mats13 = yh[1:14]
def curve(r): return r[1:14]
def avgw(a, b):
    seg = [curve(r) for r in (full[a:b] if b != 0 else full[a:])]
    return [round(sum(c)/len(c), 4) for c in zip(*seg)]
M = 21
avg_sets = [('1개월 평균', avgw(-M, 0)), ('1-2개월 평균', avgw(-2*M, -M)), ('2-3개월 평균', avgw(-3*M, -2*M)),
            ('3-4개월 평균', avgw(-4*M, -3*M)), ('4-5개월 평균', avgw(-5*M, -4*M)), ('5-6개월 평균', avgw(-6*M, -5*M))]
days5 = [(('Today' if i == 0 else f'Today -{i}') + f' ({full[-1-i][0]})', curve(full[-1-i])) for i in range(5)]
AVG_COLORS = ['#f4a261','#e9c46a','#8ab17d','#6a994e','#577590','#9d8189']
DAY_COLORS = ['#d62828','#4da3ff','#8b98a5','#577590','#3a4a5c']
def ds(label, data, color, width=1.3):
    return {'label': label, 'data': data, 'borderColor': color, 'backgroundColor': color,
            'borderWidth': width, 'pointRadius': 2, 'tension': 0.3}
c1_ds = [ds(f'Today ({full[-1][0]})', curve(full[-1]), '#d62828', 2.2)] + [ds(n, v, c) for (n, v), c in zip(avg_sets, AVG_COLORS)]
c2_ds = [ds(n, v, DAY_COLORS[i], 2.2 if i == 0 else 1.3) for i, (n, v) in enumerate(days5)]
yc_head = ''.join(f'<th>{esc(h)}</th>' for h in yh[:14])
yc_vals = ''.join(f'<td>{"" if v is None else f"{v:.2f}"}</td>' for v in yd_last[1:14])
yd_body = f"""{CHART_HDR}
<div class="card"><h2>Today 기준 평균(월)</h2><div style="height:360px;position:relative"><canvas id="c1"></canvas></div>
<div class="muted">Yield_Chart 시트 차트 1과 동일 — Today + 1~6개월 구간 평균 곡선</div></div>
<div class="card"><h2>Today ~ Today-4</h2><div style="height:360px;position:relative"><canvas id="c2"></canvas></div>
<div class="muted">Yield_Chart 시트 차트 2와 동일 — 최근 5영업일 곡선</div></div>
<div class="card"><h2>금리 추이 (10Y · 2Y · 30Y · 10Y2Y)</h2><div style="height:340px;position:relative"><canvas id="c3"></canvas></div></div>
<div class="card"><h2>최신 수치 ({yd_last[0]})</h2>
<div style="overflow-x:auto"><table class="snap"><tr><th></th>{yc_head}</tr><tr><td>{yd_last[0]}</td>{yc_vals}</tr></table></div></div>
<script>const Y={json.dumps(yd)};{JSBASE}
const mats={json.dumps(mats13)};
new Chart(c1,{{type:'line',data:{{labels:mats,datasets:{json.dumps(c1_ds, ensure_ascii=False)}}},options:{{maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}}}}}});
new Chart(c2,{{type:'line',data:{{labels:mats,datasets:{json.dumps(c2_ds, ensure_ascii=False)}}},options:{{maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}}}}}});
const L=Y.map(r=>r[0]);
new Chart(c3,{{type:'line',data:{{labels:L,datasets:[{{label:'10Y',data:Y.map(r=>r[11]),borderColor:'#4da3ff',pointRadius:0,borderWidth:1.6}},{{label:'2Y',data:Y.map(r=>r[7]),borderColor:'#3fb950',pointRadius:0,borderWidth:1.3}},{{label:'30Y',data:Y.map(r=>r[13]),borderColor:'#bc8cff',pointRadius:0,borderWidth:1.3}},{{label:'10Y2Y(우)',data:Y.map(r=>r[14]),borderColor:'#d29922',pointRadius:0,borderWidth:1.3,yAxisID:'y2'}}]}},options:{{maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},scales:{{x:{{ticks:{{maxTicksLimit:10}}}},y2:{{position:'right',grid:{{drawOnChartArea:false}}}}}}}}}});
</script>"""
open(os.path.join(SITE,'yield.html'),'w',encoding='utf-8').write(page('Yield','yield', yd_body))
# index
latest_lbl, latest_html = b26[-1]
lg_small = ''.join(f'<a href="{u}" target="_blank"><div class="t">{esc(t)}</div><div class="s">{esc(s)}</div></a>' for t,u,s in LINKS6)
idx = f"""<div class="card week"><h2 style="font-size:18px">최신 주간 브리핑 — {esc(latest_lbl)}</h2>{latest_html}</div>
<div class="card"><h2>퀵 링크</h2><div class="linkgrid">{lg_small}</div></div>
<p class="muted">과거 주차: <a href="archive-2026.html" style="color:var(--accent)">2026</a> · <a href="archive-2025.html" style="color:var(--accent)">2025</a> · <a href="archive-2024.html" style="color:var(--accent)">2023–2024</a></p>"""
open(os.path.join(SITE,'index.html'),'w',encoding='utf-8').write(page('최신 브리핑','index', idx))
print('site built OK (privacy filter active)')
