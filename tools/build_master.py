# -*- coding: utf-8 -*-
"""入試マスタ（master.xlsx）を作る。
usage: python3 -I build_master.py <raw.json> <haiten.json> <out.xlsx>
"""
import sys, json, re, hashlib, collections
sys.path.insert(0, __import__('os').path.dirname(__file__))
from rules import general, extra_general, kyotest, GROUPS, FAC_FILTER, DEPT_FILTER, D, R, ABBR
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

raw, haiten_p, out = sys.argv[1:4]
GROUP_OF = {u: g for g, us in GROUPS.items() for u in us}
METHOD_HINT = ('方式', '日程', '全学', '学部', '統一', '英語', '技能', '科目', '教科', '型', 'コア', 'プラス', '期',
               '共通', '併用', '利用', '地域枠', 'ＴＥＡＰ', '多面的')

rows = []
def add(**k):
    rows.append(k)

def clean_rank(s):
    s = (s or '').replace('\n', ' ').strip()
    if '設定' in s:
        return '設定なし'
    m = re.match(r'([\d.]+～[\d.]+)', s)
    return m.group(1) if m else s

# ---------- 1. Kei-Net 入試日一覧 ----------
seen = set()
unmatched = collections.Counter()
for x in json.load(open(raw)):
    u, rest, date = x['uni'], x['rest'], x['date']
    note = []
    free = '[自]' in rest
    mm = re.search(r'<(.+?)>', rest)
    if mm:
        note.append(mm.group(1))
    rest = re.sub(r'<.+?>', '', rest).replace('[自]', '').replace('[指]', '').strip()
    p = rest.split('-')
    if len(p) >= 3:
        f, g, m = p[0], p[1], '-'.join(p[2:])
    elif len(p) == 2:
        if any(h in p[1] for h in METHOD_HINT):
            f, g, m = p[0], '', p[1]
        else:
            f, g, m = p[0], p[1], '一般'
    else:
        f, g, m = p[0], '', '一般'
    kind = '共テ併用' if '(テ)' in m else '一般'
    m = m.replace('(テ)', '')
    key = (u, f, g, m, date)
    if key in seen:
        continue
    seen.add(key)
    if free:
        note.insert(0, '試験日自由選択（他日程でも受験可）')
    r = general(u, f, g, m, date)
    if r is None:
        unmatched[(u, f, m)] += 1
        r = R('', '', '', '', conf='要確認')
    if r['note']:
        note.append(r['note'])
    add(u=u, f=f, g=g, m=m, kind=kind, date=date, rank=clean_rank(x['rank']), r=r, note='／'.join(note), src2='Kei-Net入試日一覧')

# ---------- 2/3. 配点データ（方式一覧）から ----------
H = json.load(open(haiten_p))
def is_kyo(r):
    c, s = r['共通テ総点'], r['二次総点']
    return c not in (None, '', '0', 0) and s in (None, '', 0, '0')

keinet_unis = {x['uni'] for x in json.load(open(raw))}
seen2 = set()
for h in H:
    u = h['大学名']
    if u not in GROUP_OF or GROUP_OF[u] == '国公立' or h['kind'] != '私':
        continue
    f, g = h['学部名'] or '', h['学科名'] or ''
    if u in FAC_FILTER and f not in FAC_FILTER[u]:
        continue
    if u in DEPT_FILTER and not any(k in g for k in DEPT_FILTER[u]):
        continue
    m = h['日程方式名'] if h['日程方式名'] not in (None, 'None') else ''
    sub = h['枝番名称'] if h['枝番名称'] not in (None, 'None') else ''
    if is_kyo(h):
        mname = (m or '共テ利用') + (f'（{sub}）' if sub else '')
        k = ('K', u, f, g, mname)
        if k in seen2:
            continue
        seen2.add(k)
        r = kyotest(u, f, g, m)
        if r is None:
            unmatched[(u, f, 'K:' + m)] += 1
            r = R('', '', '', '', conf='要確認')
        add(u=u, f=f, g=g, m=mname, kind='共テ利用', date='', rank='', r=r, note=r['note'], src2='Kei-Net科目・配点')
    elif u not in keinet_unis:
        mname = (m or '一般') + (f'（{sub}）' if sub else '')
        if '地域枠' in mname:
            continue
        k = ('G', u, f, g, mname)
        if k in seen2:
            continue
        seen2.add(k)
        res = extra_general(u, f, g, m, sub)
        if res is None:
            unmatched[(u, f, 'G:' + m)] += 1
            continue
        kind = '共テ併用' if '共通テスト併用' in ''.join(rr['note'] for _, rr in res) else '一般'
        for d, r in res:
            add(u=u, f=f, g=g, m=mname, kind=kind, date=D(d), rank='', r=r, note=r['note'], src2='Kei-Net科目・配点')

# ---------- 4. 国公立 ----------
seen3 = set()
for h in H:
    u = h['大学名']
    if GROUP_OF.get(u) != '国公立' or h['kind'] != '国':
        continue
    f, g, m = h['学部名'] or '', h['学科名'] or '', h['日程方式名'] or ''
    if m not in ('前期', '後期'):
        continue
    k = (u, f, g, m)
    if k in seen3:
        continue
    seen3.add(k)
    date = D('2/25') if m == '前期' else D('3/12')
    r = R('1/25', '2/3', '', '', note=('前期は2/25から（2日間の学部あり）。合格発表は例年3/6〜3/10頃' if m == '前期'
                                      else '後期は3/12から。合格発表は例年3/20〜3/23頃') + '。出願期間・発表日は募集要項（11月頃公表）で確認',
          conf='目安', src='国公立大学共通日程')
    add(u=u, f=f, g=g, m=m + '日程', kind='国公立', date=date, rank='', r=r, note=r['note'], src2='Kei-Net科目・配点')

# ---------- ID・並べ替え ----------
order_g = list(GROUPS)
order_u = [u for us in GROUPS.values() for u in us]
rows.sort(key=lambda z: (order_g.index(GROUP_OF[z['u']]), order_u.index(z['u']), z['kind'] == '共テ利用', z['f'], z['g'], z['date'] or '9', z['m']))
ids = set()
for z in rows:
    base = '|'.join([z['u'], z['f'], z['g'], z['m'], z['kind'], z['date']])
    i = hashlib.md5(base.encode()).hexdigest()[:8]
    while i in ids:
        i = hashlib.md5((base + i).encode()).hexdigest()[:8]
    ids.add(i)
    z['id'] = i

# ---------- 書き出し ----------
wb = Workbook()
ws = wb.active
ws.title = '入試日程'
cols = ['ID', '区分', '大学', '学部', '学科', '方式', '種別', '試験日', '出願開始', '出願締切', '合格発表', '手続締切',
        '偏差値帯', '備考', '確度', '出典']
ws.append(cols)
for z in rows:
    r = z['r']
    ws.append([z['id'], GROUP_OF[z['u']], z['u'], z['f'], z['g'], z['m'], z['kind'], z['date'], r['open'], r['close'],
               r['ann'], r['proc'], z['rank'], z['note'], r['conf'], r['src'] or z['src2']])
widths = [10, 12, 14, 18, 26, 22, 9, 11, 11, 11, 11, 11, 12, 50, 14, 24]
hdr = PatternFill('solid', fgColor='1F3A5F')
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
    c = ws.cell(row=1, column=i)
    c.font = Font(bold=True, color='FFFFFF')
    c.fill = hdr
ws.freeze_panes = 'D2'
ws.auto_filter.ref = ws.dimensions

w2 = wb.create_sheet('大学')
w2.append(['大学', '表示名', '略称', '区分', '国私'])
for g, us in GROUPS.items():
    for u in us:
        disp = u if u.endswith('大学') else (u + '大学' if u not in ('慶應義塾',) else '慶應義塾大学')
        if u == '東京慈恵会医科':
            disp = '東京慈恵会医科大学'
        w2.append([u, disp, ABBR.get(u, u), g, '国公立' if g == '国公立' else '私立'])
for i, w in enumerate([16, 22, 10, 16, 8], 1):
    w2.column_dimensions[get_column_letter(i)].width = w
    w2.cell(row=1, column=i).font = Font(bold=True)

w3 = wb.create_sheet('設定')
w3.append(['項目', '値', '説明'])
w3.append(['年度', '2027', '入学年度（2027年4月入学）'])
w3.append(['共通テスト1日目', '2027-01-16', ''])
w3.append(['共通テスト2日目', '2027-01-17', ''])
w3.append(['カレンダー開始', '2026-12-01', '受験カレンダーに表示する最初の月'])
w3.append(['カレンダー終了', '2027-03-31', ''])
w3.append(['データ更新日', '2026-10-16', '画面に表示されます'])
w3.append(['注意文', '日程は2026年10月時点の公表情報です。出願前に必ず各大学の募集要項で確認してください。', '画面上部に表示'])
for i, w in enumerate([18, 60, 30], 1):
    w3.column_dimensions[get_column_letter(i)].width = w
    w3.cell(row=1, column=i).font = Font(bold=True)

w4 = wb.create_sheet('使い方')
for line in [
    'このファイルが受験校決定ページの大学データです（data/master.xlsx）。',
    '「入試日程」シート：1行＝1つの試験（試験日ごと）。共テ利用・国公立も同じ表に入っています。',
    '・日付は 2027-02-14 の形式（Excelの日付型でも可）。',
    '・ID は生徒の登録と結びつくキーです。既存行の ID は変更しないでください。新しい行は ID を空欄にすると自動で振られます。',
    '・種別は「一般」「共テ併用」「共テ利用」「国公立」のいずれか。共テ利用は試験日を空欄にします。',
    '・確度：2027確定／目安／前年参考／要確認。画面で色分け表示されます。',
    '「大学」シート：大学の並び順と区分（早慶上理など）。ここにない大学も「入試日程」に行があれば表示されます。',
    '「設定」シート：年度・共通テスト日・カレンダー期間など。',
    '次年度の更新：教員画面の「データ更新」でKei-Netの入試日一覧xlsxを読み込むと、この形式の行が作れます。',
]:
    w4.append([line])
w4.column_dimensions['A'].width = 120
wb.save(out)

print('rows', len(rows))
print(collections.Counter(z['kind'] for z in rows))
print(collections.Counter(z['r']['conf'] for z in rows))
print('unmatched:')
for k, n in unmatched.most_common(80):
    print(' ', k, n)
