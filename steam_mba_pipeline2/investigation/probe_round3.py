"""FASE 1c - Confirmacoes criticas.

1) formato de store/actions/GetAppList
2) IDs REAIS das tags (4166 = Atmospheric, nao 3D!)
3) age gate real (jogos com restricao efetiva)
4) limites de paginacao do /search/results/
5) data-ds-tagids dentro do search HTML (tags de graca no discovery?)
6) campo recommendations no appdetails (review count barato)
"""
from __future__ import annotations

import json
import os
import re
import time

import requests

HERE = os.path.dirname(__file__)
OUT = os.path.join(HERE, "samples")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36 (academic-research)")
S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})


def get(url, **kw):
    kw.setdefault("timeout", 30)
    r = S.get(url, **kw)
    time.sleep(1.2)
    return r


print("=" * 78)
print("1) store/actions/GetAppList/?appType=game  -- qual o formato?")
print("=" * 78)
r = get("https://store.steampowered.com/actions/GetAppList/", params={"appType": "game"})
print(f"  HTTP {r.status_code} bytes={len(r.content)} ct={r.headers.get('Content-Type')}")
print(f"  primeiros 300 chars: {r.text[:300]!r}")
try:
    j = r.json()
    print(f"  json OK, tipo={type(j).__name__}, len={len(j)}")
    print(f"  amostra: {json.dumps(j[:3] if isinstance(j, list) else j, ensure_ascii=False)[:300]}")
    with open(os.path.join(OUT, "actions_getapplist_game.json"), "w", encoding="utf-8") as fh:
        json.dump(j, fh, ensure_ascii=False)
except Exception as e:  # noqa: BLE001
    print(f"  json FALHOU: {e}")
    with open(os.path.join(OUT, "actions_getapplist_game.raw"), "w",
              encoding="utf-8", errors="replace") as fh:
        fh.write(r.text)

print()
print("=" * 78)
print("2) TAG IDS REAIS - endpoints candidatos")
print("=" * 78)
tag_endpoints = [
    ("ajaxgetstoretags", "https://store.steampowered.com/actions/ajaxgetstoretags", None),
    ("IStoreService/GetTagList (sem key)",
     "https://api.steampowered.com/IStoreService/GetTagList/v1/", {"language": "english"}),
    ("tagdata/populartags",
     "https://store.steampowered.com/tagdata/populartags/english", None),
]
tag_map: dict[str, int] = {}
for label, url, params in tag_endpoints:
    try:
        r = get(url, params=params)
        print(f"  {label:38s} HTTP {r.status_code} bytes={len(r.content)}")
        if r.status_code == 200:
            try:
                j = r.json()
                items = j if isinstance(j, list) else (
                    j.get("tags") or j.get("response", {}).get("tags") or [])
                print(f"     -> {len(items)} tags. amostra: "
                      f"{json.dumps(items[:5], ensure_ascii=False)[:250]}")
                for it in items:
                    nm = it.get("name")
                    tid = it.get("tagid") or it.get("tagId")
                    if nm and tid:
                        tag_map[nm] = tid
                with open(os.path.join(OUT, f"tags_{label.split('/')[0].split(' ')[0]}.json"),
                          "w", encoding="utf-8") as fh:
                    json.dump(j, fh, ensure_ascii=False, indent=2)
            except Exception as e:  # noqa: BLE001
                print(f"     json falhou: {e}")
    except Exception as e:  # noqa: BLE001
        print(f"  {label:38s} ERRO {type(e).__name__}: {e}")

WANTED = ["3D", "2D", "2.5D", "Indie", "First-Person", "Third Person",
          "Third-Person Shooter", "FPS", "3D Platformer", "3D Fighter",
          "Immersive Sim", "Walking Simulator", "Automobile Sim", "Flight",
          "Space Sim", "Looter Shooter", "Hero Shooter", "Arena Shooter",
          "Rail Shooter", "Open World", "Realistic", "Cinematic",
          "Action-Adventure", "Action RPG", "Souls-like", "Survival Horror",
          "Driving", "2D Platformer", "2D Fighter", "Pixel Graphics",
          "Hand-drawn", "Side Scroller", "Text-Based", "Visual Novel",
          "Interactive Fiction", "Point & Click", "Card Game", "Board Game",
          "Hidden Object", "Casual", "Free to Play", "Early Access", "Homemade"]
print()
print("  RESOLUCAO DAS TAGS DO PROJETO (nome -> tagid):")
if tag_map:
    print(f"  (mapa carregado com {len(tag_map)} tags)")
    for w in WANTED:
        print(f"    {w:24s} -> {tag_map.get(w, 'NAO ENCONTRADA')}")
else:
    print("  NENHUM endpoint de tags funcionou -> tags terao de vir do HTML por app")

print()
print("=" * 78)
print("3) AGE GATE real")
print("=" * 78)
for label, appid in [("Postal2_AO", 232770), ("HuniePop", 339800),
                     ("Manhunt", 12130), ("RDR2", 1174180)]:
    try:
        r = S.get(f"https://store.steampowered.com/app/{appid}/",
                  timeout=30, allow_redirects=True)
        low = r.text.lower()
        print(f"  {label:12s} HTTP {r.status_code} final={r.url[:70]}")
        print(f"     agecheck={'agecheck' in low} "
              f"sys_req={'game_area_sys_req' in r.text} "
              f"tagmodal={'InitAppTagModal' in r.text} bytes={len(r.content)}")
        time.sleep(1.2)
    except Exception as e:  # noqa: BLE001
        print(f"  {label:12s} ERRO {e}")

print()
print("=" * 78)
print("4) PAGINACAO PROFUNDA do /search/results/")
print("=" * 78)
for start in (0, 1000, 5000, 20000, 38000, 45000):
    r = get("https://store.steampowered.com/search/results/",
            params={"query": "", "start": start, "count": 50, "infinite": 1,
                    "category1": 998, "os": "win", "json": 1})
    try:
        j = r.json()
        ids = re.findall(r'data-ds-appid="([\d,]+)"', j.get("results_html", ""))
        print(f"  start={start:>6} HTTP {r.status_code} total_count={j.get('total_count')} "
              f"retornados={len(ids)}")
    except Exception as e:  # noqa: BLE001
        print(f"  start={start:>6} HTTP {r.status_code} json falhou: {e}")

print()
print("=" * 78)
print("5) data-ds-tagids no HTML do search? (tags gratis no discovery)")
print("=" * 78)
r = get("https://store.steampowered.com/search/results/",
        params={"query": "", "start": 0, "count": 25, "infinite": 1,
                "category1": 998, "os": "win", "json": 1,
                "sort_by": "Released_DESC"})
html = r.json().get("results_html", "")
with open(os.path.join(OUT, "search_row_sample.html"), "w", encoding="utf-8",
          errors="replace") as fh:
    fh.write(html)
row = re.search(r'<a[^>]*search_result_row.*?</a>', html, re.S)
if row:
    attrs = dict(re.findall(r'(data-ds-[a-z]+)="([^"]*)"', row.group(0)))
    print(f"  atributos data-ds-* da 1a linha: {json.dumps(attrs, ensure_ascii=False)[:500]}")
    txt = re.sub(r"<[^>]+>", " | ", row.group(0))
    print(f"  texto da linha: {re.sub(r'\\s+', ' ', txt)[:400]}")
print(f"  ocorrencias de data-ds-tagids no lote: {html.count('data-ds-tagids')}")

print()
print("=" * 78)
print("6) review count barato: campo 'recommendations' do appdetails")
print("=" * 78)
import glob
for p in sorted(glob.glob(os.path.join(OUT, "appdetails_*.json"))):
    with open(p, encoding="utf-8") as fh:
        j = json.load(fh)
    node = next(iter(j.values()))
    if node.get("success"):
        d = node["data"]
        print(f"  {os.path.basename(p):46s} recommendations={d.get('recommendations')} "
              f"metacritic={(d.get('metacritic') or {}).get('score')}")
