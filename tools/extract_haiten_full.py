# -*- coding: utf-8 -*-
"""Kei-Net「科目・配点データ」（私立・国公立のxlsx）から、対象大学の行だけを全列で取り出す。
usage: python3 -I extract_haiten_full.py <master.xlsx（大学シートの大学が対象）> <out.json> <私立.xlsx> <国公立.xlsx>
出力した json を build_master.py の4番目の引数に渡すと、受験科目・配点の列が埋まる。
"""
import sys, json, openpyxl

master, out_p, shi, koku = sys.argv[1:5]
T = set(r[0] for r in openpyxl.load_workbook(master)['大学'].iter_rows(min_row=2, values_only=True) if r[0])
out = []
for f, kind in [(shi, '私'), (koku, '国')]:
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True).active
    it = ws.iter_rows(values_only=True)
    h = next(it)
    for r in it:
        d = {h[i]: v for i, v in enumerate(r) if v not in (None, '')}
        if d.get('大学名') in T:
            d['kind'] = kind
            out.append(d)
json.dump(out, open(out_p, 'w'), ensure_ascii=False)
print(len(out), 'rows')
