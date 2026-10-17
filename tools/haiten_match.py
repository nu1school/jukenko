# -*- coding: utf-8 -*-
"""入試日程の1行（大学・学部・学科・方式・種別）に対応する Kei-Net「科目・配点」の行を探す。"""
import collections, re
D = []
def nz(s): 
    s=(s or '').replace('None','')
    return re.sub(r'[\s　・]','',s)
def norm_m(m):
    m=nz(m)
    if m=='一般': m=''
    m=re.sub(r'（.*?）|\(.*?\)','',m)
    return m.replace('日程','').replace('方式','').replace('入試','').replace('試験','')
idx=collections.defaultdict(list)
def load(rows):
    D.extend(rows)
    for d in rows: idx[(d['大学名'],nz(d.get('学部名')),nz(d.get('学科名')))].append(d)
def is_kyo_only(d): return bool(d.get('共通テ総点')) and not d.get('二次総点')
def has_niji(d): return bool(d.get('二次総点'))
def sim(a,b):
    if not a and not b: return 1
    if not a or not b: return 0.2
    A=set(a);B=set(b); return len(A&B)/len(A|B) + (0.5 if a in b or b in a else 0)
def fmatch(a,b):
    a=re.sub(r'（.*?）','',nz(a)); b=nz(b).split('－')[0]
    return a==b or (len(b)>=3 and a.startswith(b)) or (len(a)>=3 and b.startswith(a))
def gmatch(a,b):
    a=nz(a); b=nz(b)
    return a==b or (a and b and (a in b or b in a)) or (a and b.split('－')[-1]==a)
def lookup(u,f,g):
    c=idx.get((u,nz(f),nz(g)))
    if c: return c
    ks=[k for k in idx if k[0]==u and fmatch(f,k[1])]
    c=[d for k in ks if nz(k[2])==nz(g) for d in idx[k]]
    if c: return c
    c=[d for k in ks if gmatch(g,k[2]) for d in idx[k]]
    if c: return c
    if not g: return [d for k in ks for d in idx[k]]
    return None
def find(u,f,g,m,kind):
    cands=lookup(u,f,g)
    if not cands: return None,'nokey'
    allc=cands
    mm0=re.match(r'(.*?)（(.*)）$',m); b0=(mm0.group(1) if mm0 else m)
    if kind=='国公立': b0=b0.replace('日程','')
    ex=[d for d in allc if norm_m(d.get('日程方式名'))==norm_m(b0) and norm_m(b0)]
    if ex:
        sub0=mm0.group(2) if mm0 else ''
        ex2=[d for d in ex if nz(d.get('枝番名称'))==nz(sub0)]
        return (ex2 or ex)[0],'exact'
    if kind=='共テ利用': cands=[d for d in cands if is_kyo_only(d)]
    elif kind=='国公立': cands=[d for d in cands if has_niji(d) or d.get('日程方式名') in ('前期','後期')]
    elif kind=='共テ併用': cands=[d for d in cands if d.get('共通テ総点') and has_niji(d)] or [d for d in cands if has_niji(d)]
    else: cands=[d for d in cands if has_niji(d) and not d.get('共通テ総点')] or [d for d in cands if has_niji(d)]
    if not cands: cands=allc
    mm=re.match(r'(.*?)（(.*)）$',m)
    base,sub=(mm.group(1),mm.group(2)) if mm else (m,'')
    if kind=='国公立': base=base.replace('日程','')
    best=None;bs=-1
    for d in cands:
        s=sim(norm_m(base),norm_m(d.get('日程方式名')))+ (0.3*sim(nz(sub),nz(d.get('枝番名称'))) if sub or d.get('枝番名称') else 0.3)
        if s>bs: bs,best=s,d
    if len(cands)==1: return best,'single'
    if bs<0.6 and cands is not allc:
        for d in allc:
            s=sim(norm_m(base),norm_m(d.get('日程方式名')))+0.3
            if s>bs: bs,best=s,d
    return (best,'sim%.1f'%bs) if bs>=0.6 else (None,'low')
