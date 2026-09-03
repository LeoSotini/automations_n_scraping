"""FASE 1d - Semantica do search, funil real por tag, e risco do data-ds-tagids."""
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

BASE = "https://store.steampowered.com/search/results/"


def count(params: dict) -> tuple[int | None, int]:
    p = {"query": "", "start": 0, "count": 1, "infinite": 1, "json": 1,
         "os": "win", "category1": 998, **params}
    r = S.get(BASE, params=p, timeout=30)
    time.sleep(1.0)
    try:
        j = r.json()
        return j.get("total_count"), r.status_code
    except Exception:
        return None, r.status_code


TAGS = {
    "3D": 4191, "First-Person": 3839, "Third Person": 1697,
    "Third-Person Shooter": 3814, "FPS": 1663, "3D Platformer": 5395,
    "3D Fighter": 6506, "Immersive Sim": 9204, "Walking Simulator": 5900,
    "Automobile Sim": 1100687, "Flight": 15045, "Space Sim": 16598,
    "Looter Shooter": 353880, "Hero Shooter": 620519, "Arena Shooter": 5547,
    "Rail Shooter": 3954,
    "Open World": 1695, "Realistic": 4175, "Cinematic": 4145,
    "Action-Adventure": 4106, "Action RPG": 4231, "Souls-like": 29482,
    "Survival Horror": 3978, "Driving": 1644,
    "2D": 3871, "2D Platformer": 5379, "2D Fighter": 4736, "2.5D": 4975,
    "Indie": 492,
}

print("=" * 78)
print("1) SEMANTICA: multiplas tags no search = AND ou OR?")
print("=" * 78)
c3d, _ = count({"tags": 4191})
cfps, _ = count({"tags": 1663})
cboth, _ = count({"tags": "4191,1663"})
print(f"  3D isolado           = {c3d:,}")
print(f"  FPS isolado          = {cfps:,}")
print(f"  '4191,1663' juntos   = {cboth:,}")
if cboth is not None and c3d and cfps:
    if cboth <= min(c3d, cfps):
        print("  -> COMPORTAMENTO = AND (intersecao). Nao serve para OR de tags positivas.")
    elif cboth >= max(c3d, cfps):
        print("  -> COMPORTAMENTO = OR (uniao).")
    else:
        print("  -> COMPORTAMENTO AMBIGUO. Investigar.")

print()
print("=" * 78)
print("2) FUNIL: total_count por tag (os=win, category1=998=jogos)")
print("=" * 78)
base_win, _ = count({})
print(f"  BASELINE jogos Windows na store = {base_win:,}\n")
results = {}
for name, tid in TAGS.items():
    c, st = count({"tags": tid})
    results[name] = c
    print(f"  {name:22s} tagid={tid:<8} -> {c:>8,}" if c is not None
          else f"  {name:22s} tagid={tid:<8} -> ERRO HTTP {st}")

print()
print("=" * 78)
print("3) EFEITO DAS EXCLUSOES sobre a tag 3D (via AND, ja que search e AND)")
print("=" * 78)
c_3d_indie, _ = count({"tags": "4191,492"})
c_3d_2d, _ = count({"tags": "4191,3871"})
print(f"  3D total              = {c3d:,}")
print(f"  3D AND Indie          = {c_3d_indie:,}   "
      f"({100*c_3d_indie/c3d:.1f}% do 3D e Indie)")
print(f"  3D AND 2D (conflito!) = {c_3d_2d:,}   "
      f"-> jogos com AMBAS as tags existem: casos ambiguos reais")

print()
print("=" * 78)
print("4) RISCO: data-ds-tagids do search vs 20 tags completas do HTML")
print("=" * 78)
r = S.get(BASE, params={"query": "", "start": 0, "count": 25, "infinite": 1,
                        "json": 1, "os": "win", "category1": 998,
                        "tags": 4191, "sort_by": "Reviews_DESC"}, timeout=30)
time.sleep(1.0)
html = r.json().get("results_html", "")
rows = re.findall(r'data-ds-appid="(\d+)"[^>]*?data-ds-tagids="(\[[^\]]*\])"', html)
if not rows:
    rows = [(m.group(1), m.group(2)) for m in re.finditer(
        r'data-ds-appid="(\d+)".*?data-ds-tagids="(\[[^\]]*\])"', html, re.S)]
print(f"  linhas com tagids: {len(rows)}")
sizes = [len(json.loads(t)) for _, t in rows]
print(f"  tamanho dos vetores de tagids: min={min(sizes)} max={max(sizes)} "
      f"media={sum(sizes)/len(sizes):.1f}")
print("  -> HTML da pagina do jogo entrega 20 tags; o search entrega apenas as acima.")

# comparacao direta em 3 jogos
for appid in [int(rows[0][0]), 1091500, 1245620]:
    rr = S.get(f"https://store.steampowered.com/app/{appid}/", timeout=30)
    time.sleep(1.2)
    m = re.search(r"InitAppTagModal\(\s*\d+\s*,\s*(\[.*?\])\s*,", rr.text, re.S)
    full = {t["tagid"] for t in json.loads(m.group(1))} if m else set()
    sub = next((set(json.loads(t)) for a, t in rows if int(a) == appid), None)
    print(f"    app {appid}: HTML={len(full)} tags | search={len(sub) if sub else 'n/a'} "
          f"| search e subconjunto do HTML? "
          f"{sub.issubset(full) if sub and full else 'n/a'}")

print()
print("=" * 78)
print("5) FILTRO DE DATA no search? (parametro de intervalo de release)")
print("=" * 78)
for label, params in [
    ("sem filtro de data", {"tags": 4191}),
    ("untags=Early Access(493)", {"tags": 4191, "untags": 493}),
    ("hidef2p=1", {"tags": 4191, "hidef2p": 1}),
]:
    c, st = count(params)
    print(f"  {label:28s} -> total_count={c}")
print("  NOTA: /search nao expoe filtro de release_date por intervalo; a data sera")
print("        aplicada apos o appdetails (que traz release_date.date).")

with open(os.path.join(OUT, "_funnel_counts.json"), "w", encoding="utf-8") as fh:
    json.dump({"baseline_windows_games": base_win, "per_tag": results},
              fh, ensure_ascii=False, indent=2)
print("\n[DONE] contagens salvas em samples/_funnel_counts.json")
