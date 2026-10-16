# -*- coding: utf-8 -*-
"""Kei-Net のダウンロードファイルを build_master.py の入力（JSON）に変換する。
usage: python3 parse_keinet.py <入試日一覧xlsxのフォルダ> <科目・配点(私立)xlsx> <科目・配点(国公立)xlsx>
出力: raw.json（入試日一覧）, haiten.json（科目・配点の方式一覧）
"""
import sys, glob, re, json, openpyxl

date_dir, shiritsu, kokkou = sys.argv[1:4]
out = []
for f in sorted(glob.glob(date_dir.rstrip('/') + '/*.xlsx')):
    ws = openpyxl.load_workbook(f, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows) if r[0] and 'ランク偏差値' in str(r[0])), None)
    if hi is None:
        continue
    hdr = [(str(h).replace('_x000D_', '').replace('\n', ' ').strip() if h else '') for h in rows[hi]]
    cur = None
    for r in rows[hi + 1:]:
        if r[0] is not None:
            cur = r[0]
        for j, c in enumerate(r[1:], 1):
            if not c:
                continue
            for s in str(c).replace('_x000D_', '').split('\n'):
                m = re.match(r'【(.+?)】(.*)', s.strip())
                if m and cur:
                    out.append(dict(date=str(cur)[:10], rank=hdr[j], uni=m.group(1), rest=m.group(2)))
json.dump(out, open('raw.json', 'w'), ensure_ascii=False)

rows = []
for f, kind in [(shiritsu, '私'), (kokkou, '国')]:
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True).active
    it = ws.iter_rows(values_only=True); h = next(it); idx = {k: i for i, k in enumerate(h)}
    keys = ['年度', '大学名', '学部名', '学科名', '日程方式名', '枝番名称', '共通テ総点', '二次総点', '１０桁コード']
    for r in it:
        rows.append({k: r[idx[k]] for k in keys} | {'kind': kind})
json.dump(rows, open('haiten.json', 'w'), ensure_ascii=False, default=str)
print('入試日一覧', len(out), '行 / 科目・配点', len(rows), '行')
