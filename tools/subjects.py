# -*- coding: utf-8 -*-
"""Kei-Net「科目・配点」1行から、受験科目の要約文・配点・絞り込み用の科目要件を作る。"""
import re

def _z(v):
    return v not in (None, '', 0, '0')

def _n(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0

ZEN = str.maketrans('０１２３４５６７８９②③', '0123456789' + '23')

# (列名, 表示名, 絞り込み用の科目(複数なら/区切り))
KYO = {
    '外国語': [('共通テ英語', '英', '英'), ('共通テ独語', '独', ''), ('共通テ仏語', '仏', ''), ('共通テ中語', '中', ''), ('共通テ韓語', '韓', '')],
    '数学': [('共通テ数学ⅠＡ', 'ⅠA', '数ⅠA'), ('共通テ数学Ⅰ', 'Ⅰ', '数ⅠA'), ('共通テ数学ⅡＢＣ', 'ⅡBC', '数ⅡBC')],
    '国語': [('共通テ国語', '国', '国')],
    '理科': [('共通テ物理基礎', '基礎', '基礎'), ('共通テ化学基礎', '基礎', '基礎'), ('共通テ生物基礎', '基礎', '基礎'), ('共通テ地学基礎', '基礎', '基礎'),
           ('共通テ物理', '物', '物'), ('共通テ化学', '化', '化'), ('共通テ生物', '生', '生'), ('共通テ地学', '地学', '地学')],
    '地公': [('共通テ地理', '地理', '地理'), ('共通テ日本史', '日', '日'), ('共通テ世界史', '世', '世'), ('共通テ地歴公共', '地総歴総公', '地理/日/世/倫/政'),
           ('共通テ倫理', '倫', '倫'), ('共通テ政経', '政', '政')],
    '情報': [('共通テ情報', '情', '情')],
}
NIJI = {
    '外国語': [('二次英語', '英', '英')],
    '数学': [('二次数学', '数', '')],
    '国語': [('二次国語', '国', '国')],
    '理科': [('二次物理', '物', '物'), ('二次化学', '化', '化'), ('二次生物', '生', '生'), ('二次地学', '地学', '地学')],
    '地公': [('二次地総', '地総', '地理'), ('二次歴総', '歴総', '日/世'), ('二次地理', '地理', '地理'), ('二次日本史', '日', '日'), ('二次世界史', '世', '世'),
           ('二次公共', '公共', '倫/政'), ('二次倫理', '倫', '倫'), ('二次政経', '政', '政')],
    '情報': [('二次情報', '情', '情')],
}
GSHORT = {'外国語': '英', '数学': '数', '国語': '国', '理科': '理', '地公': '地公', '情報': '情'}
PAT_GROUP = {'英': '外国語', '外': '外国語', '数': '数学', '国': '国語', '理': '理科', '地公': '地公', '地歴': '地公', '公民': '地公', '情': '情報'}
PAT_ITEM = {'倫': '倫', '政': '政', '地理': '地理', '日': '日', '世': '世', '物': '物', '化': '化', '生': '生'}
EXTRAS = [('英資', '英資'), ('小論文', '小論'), ('総合', '総合'), ('面接', '面接'), ('実技', '実技'), ('簿記', '簿記'),
          ('理数', '理数'), ('調査書', '調査書'), ('書類', '書類'), ('他', 'その他')]


def _niji_math(d):
    f = [k for k, col in (('Ⅰ', '数学１'), ('Ⅱ', '数学２'), ('Ⅲ', '数学３'), ('A', '数学Ａ'), ('B', '数学Ｂ'), ('C', '数学Ｃ')) if _z(d.get(col))]
    lab = '数' + ''.join(f) if f else '数'
    tag = '数Ⅲ' if 'Ⅲ' in f else ('数ⅡBC' if ('Ⅱ' in f or 'B' in f or 'C' in f) else '数ⅠA')
    return lab, tag


def _pts(d, pre, g):
    if g == '情報':
        return _n(d.get(pre + '情報－配点')) if pre == '共通テ' else _n(d.get('二次情報－配点小'))
    return _n(d.get(f'{pre}{g}－配点小'))


def _parse_pattern(s):
    """'＊数・理・地公・情→１、英資・他→１' -> [(['数学','理科',...], ['倫',...], 1), ...]"""
    out = []
    if not s:
        return out
    s = s.lstrip('＊*').translate(ZEN)
    for cl in re.split('[、,]', s):
        if '→' not in cl:
            continue
        left, right = cl.rsplit('→', 1)
        m = re.match(r'\d+', right.strip())
        if not m:
            continue
        if '必須' in left:
            left = left.split('必須')[-1]
        groups, items, extras = [], [], []
        for t in left.split('・'):
            t = re.sub(r'（.*?）', '', t).strip()
            t = re.sub(r'\d+$', '', t)
            if t in PAT_GROUP:
                groups.append(PAT_GROUP[t])
            elif t in PAT_ITEM:
                items.append(PAT_ITEM[t])
            elif t:
                extras.append(t)
        out.append((groups, items, extras, int(m.group())))
    return out


def stage(d, pre):
    """1段階（共通テ or 二次）の要約。戻り値: (表示文, 総点, slots)  slots=[(alternatives set, count)]"""
    total = _n(d.get(pre + '総点'))
    T = KYO if pre == '共通テ' else NIJI
    parts, slots = [], []
    pool = {}  # group -> (labels, tags, pts)
    for g, cols in T.items():
        req, opt, pl = [], [], []
        rtag, otag, ptag = [], [], []
        for col, lab, tag in cols:
            v = d.get(col)
            if not v:
                continue
            if g == '数学' and pre == '二次':
                lab, tag = _niji_math(d)
            mark = v[0]
            if pre == '二次' and g == '理科' and '基' in (d.get(col + '範囲') or ''):
                lab = lab + '基'
            tags = [t for t in tag.split('/') if t]
            if mark == '●':
                if lab not in req: req.append(lab)
                rtag.append(tags)
            elif mark == '○':
                if lab not in opt: opt.append(lab)
                otag += tags
            elif mark == '◇':
                if lab not in pl: pl.append(lab)
                ptag += tags
        if not (req or opt or pl):
            continue
        p = _pts(d, pre, g)
        if g == '外国語':
            # 英語が●なら他言語の○は無視（英語で受ける前提）
            if '英' in req:
                opt = []
                otag = []
            elif '英' in opt:
                req, opt, rtag, otag = ['英'], [], [['英']], []
            lab = '英'
            if pre == '共通テ':
                r, l = _n(d.get('共通テ英語－配点')), _n(d.get('共通テリスニング－配点'))
                if r and l and r != l:
                    lab = f'英(R{r}L{l})'
                elif r and not l and _z(d.get('共通テリスニング')) is False:
                    lab = '英(Rのみ)'
            if req or opt:
                parts.append(f'{lab}{p or ""}')
                slots.append(({'英'}, 1))
            if pl:
                pool[g] = (['英'], {'英'}, p)
            continue
        nsub = _n(d.get(f'{pre}{g}－科目数')) if g != '情報' else 1
        txt = ''
        if g == '国語':
            rng = d.get(pre + '現古漢範囲') or ''
            base = '国' + (f'({rng})' if rng and rng not in ('その他',) else '')
            if req or opt:
                parts.append(f'{base}{p or ""}')
                slots.append(({'国'}, 1))
            if pl:
                pool[g] = ([base], {'国'}, p)
            continue
        if req:
            if g == '数学':
                txt = '数' + '・'.join(x.replace('数', '') for x in req) if pre == '共通テ' else '・'.join(req)
            else:
                txt = '・'.join(req)
            for t in rtag:
                slots.append((set(t), 1))
        if opt:
            k = max(1, nsub - len(req)) if nsub else 1
            names = opt if not (g == '数学' and pre == '共通テ') else ['数' + x for x in opt]
            o = '〔' + '/'.join(names) + (f'から{k}' if k > 1 else '') + '〕'
            txt = (txt + '＋' if txt else '') + o
            slots.append((set(otag), k))
        if txt:
            parts.append(f'{txt}{p or ""}')
        if pl:
            names = pl if not (g == '数学' and pre == '共通テ') else ['数' + x for x in pl]
            pool[g] = (names, set(ptag), p)
    # 外国語以外の付帯（二次の英資・小論文・面接など）
    extras = []
    if pre == '二次':
        for key, lab in EXTRAS:
            v = d.get('二次' + key)
            if not v:
                continue
            p = _n(d.get(f'二次{key}－配点'))
            mark = v[0]
            if key == '英資':
                s = '英語資格' + ('必須' if mark == '●' else '可')
            else:
                s = lab
            if mark == '◇':
                pool[key] = ([s], set(), p)
                continue
            extras.append(s + (str(p) if p else ''))
    else:
        v = d.get('共通テ英資')
        if v:
            p = _n(d.get('共通テ英資－配点'))
            extras.append('英語資格' + ('必須' if v[0] == '●' else '可') + (str(p) if p else ''))
    # 教科をまたぐ選択
    if pool:
        pats = _parse_pattern(d.get(pre + '教科選択パターン'))
        used = set()
        for groups, items, ex, k in pats:
            names, tags, pts = [], set(), 0
            for g in groups:
                if g in pool:
                    names.append(GSHORT[g]); tags |= pool[g][1]; pts = pts or pool[g][2]; used.add(g)
            for it in items:
                for g, (ns, ts, pp) in pool.items():
                    if it in ts:
                        names.append(it); tags.add(it); pts = pts or pp; used.add(g)
            for e in ex:
                for g in pool:
                    if g.startswith(e) or e.startswith(g[:2]):
                        names.append(pool[g][0][0]); pts = pts or pool[g][2]; used.add(g)
                        tags.add('*')  # 学力以外の選択肢 → 科目による制約なし
            if names:
                parts.append(f'〈{"/".join(dict.fromkeys(names))}{f"から{k}" if k > 1 else ""}〉{pts or ""}')
                slots.append((tags, k))
        rest = [g for g in pool if g not in used]
        if rest:
            names = [GSHORT.get(g, pool[g][0][0]) for g in rest]
            tags = set().union(*[pool[g][1] for g in rest])
            if any(not pool[g][1] for g in rest):
                tags.add('*')
            parts.append(f'〈{"/".join(names)}〉{pool[rest[0]][2] or ""}')
            slots.append((tags, 1))
    parts += extras
    return '・'.join(parts), total, slots


def slots_str(slots):
    """絞り込み用：'英 国 数ⅠA/物/化*1' （空白区切りの枠、/は選択肢、*nは必要数）"""
    out = []
    for alts, k in slots:
        if not alts:
            continue
        a = '/'.join(sorted(alts, key=ORDER.index if all(x in ORDER for x in alts) else str))
        out.append(a + (f'*{k}' if k != 1 else ''))
    return ' '.join(out)

ORDER = ['英', '国', '数ⅠA', '数ⅡBC', '数Ⅲ', '物', '化', '生', '地学', '基礎', '日', '世', '地理', '倫', '政', '情', '*']


def summarize(d):
    """配点の1行 → dict(共テ科目, 共テ配点, 個別科目, 個別配点, 科目要件)"""
    out = {'kt': '', 'kp': '', 'nt': '', 'np': '', 'req': ''}
    reqs = []
    if _n(d.get('共通テ総点')):
        t, p, s = stage(d, '共通テ')
        out['kt'], out['kp'] = t, p
        reqs.append(('テ', s))
    if _n(d.get('二次総点')) or any(d.get('二次' + k) for k, _ in EXTRAS) or d.get('二次英語'):
        t, p, s = stage(d, '二次')
        out['nt'], out['np'] = t, (p or '')
        reqs.append(('個', s))
    if d.get('英資出願要件') == '★':
        out['nt'] = ('英語資格が出願要件・' + out['nt']).rstrip('・')
    out['req'] = ' | '.join(f'{k}:{slots_str(s)}' for k, s in reqs if slots_str(s))
    return out
