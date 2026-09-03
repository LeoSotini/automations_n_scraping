"""FASE 1e - Completa as medicoes que sofreram 429, agora com backoff real.

Tambem mede empiricamente o limite de taxa observado.
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
BASE = "https://store.steampowered.com/search/results/"

PAUSE = 3.0


def count(params: dict, max_retries: int = 5) -> int | None:
    p = {"query": "", "start": 0, "count": 1, "infinite": 1, "json": 1,
         "os": "win", "category1": 998, **params}
    delay = PAUSE
    for attempt in range(1, max_retries + 1):
        r = S.get(BASE, params=p, timeout=30)
        if r.status_code == 429:
            wait = delay * (2 ** (attempt - 1))
            print(f"      [429] tentativa {attempt}, aguardando {wait:.0f}s")
            time.sleep(wait)
            continue
        time.sleep(PAUSE)
        try:
            return r.json().get("total_count")
        except Exception:
            return None
    return None


print("=" * 78)
print("A) Completando o funil (tags que sofreram 429)")
print("=" * 78)
REMAINING = {"2D Fighter": 4736, "2.5D": 4975, "Indie": 492,
             "Pixel Graphics": 3964, "Side Scroller": 3798,
             "Visual Novel": 3799, "Casual": 597}
extra = {}
for name, tid in REMAINING.items():
    c = count({"tags": tid})
    extra[name] = c
    print(f"  {name:20s} tagid={tid:<8} -> {c:>8,}" if c is not None
          else f"  {name:20s} tagid={tid:<8} -> FALHOU")

print()
print("=" * 78)
print("B) Intersecoes relevantes para a metodologia (search e AND)")
print("=" * 78)
INTER = {
    "3D": {"tags": 4191},
    "3D AND Indie": {"tags": "4191,492"},
    "3D AND 2D (conflito)": {"tags": "4191,3871"},
    "3D AND 2.5D": {"tags": "4191,4975"},
    "3D AND Pixel Graphics": {"tags": "4191,3964"},
    "3D NOT Indie (untags)": {"tags": 4191, "untags": 492},
    "3D NOT Indie NOT 2D": {"tags": 4191, "untags": "492,3871"},
}
inter = {}
for label, params in INTER.items():
    c = count(params)
    inter[label] = c
    print(f"  {label:26s} -> {c:>8,}" if c is not None
          else f"  {label:26s} -> FALHOU")

print()
print("=" * 78)
print("C) 'untags' funciona como exclusao? (validacao logica)")
print("=" * 78)
t3d = inter.get("3D")
t3d_indie = inter.get("3D AND Indie")
t3d_not_indie = inter.get("3D NOT Indie (untags)")
if all(v is not None for v in (t3d, t3d_indie, t3d_not_indie)):
    soma = t3d_indie + t3d_not_indie
    print(f"  3D={t3d:,} | 3D AND Indie={t3d_indie:,} | 3D NOT Indie={t3d_not_indie:,}")
    print(f"  soma das partes={soma:,} vs total={t3d:,} | diferenca={soma - t3d:,}")
    print("  -> untags CONFIRMADO como complemento exato." if soma == t3d
          else "  -> untags NAO e complemento exato; usar com cautela.")

print()
print("=" * 78)
print("D) LIMITE DE TAXA EMPIRICO no /search/results/ (rajada controlada)")
print("=" * 78)
S2 = requests.Session()
S2.headers.update({"User-Agent": UA})
ok = 0
first429 = None
t0 = time.time()
for i in range(1, 41):
    r = S2.get(BASE, params={"query": "", "start": i * 50, "count": 1,
                             "infinite": 1, "json": 1, "os": "win"}, timeout=30)
    if r.status_code == 429:
        first429 = i
        print(f"  primeiro 429 na requisicao #{i} apos {time.time()-t0:.1f}s "
              f"({ok} respostas 200)")
        ra = r.headers.get("Retry-After")
        print(f"  Retry-After={ra!r} | headers de rate: "
              f"{[k for k in r.headers if 'rate' in k.lower()]}")
        break
    ok += 1
    time.sleep(0.35)
if first429 is None:
    print(f"  nenhum 429 em 40 requisicoes a ~0.35s ({time.time()-t0:.1f}s total)")

print()
print("=" * 78)
print("E) LIMITE DE TAXA no appdetails (endpoint principal do scraping)")
print("=" * 78)
S3 = requests.Session()
S3.headers.update({"User-Agent": UA})
appids = [730, 620, 220, 400, 440, 570, 550, 240, 300, 320,
          360, 380, 420, 500, 520, 540, 560, 580, 600, 640,
          660, 680, 700, 720, 740, 760, 780, 800, 820, 840]
ok = 0
t0 = time.time()
for i, aid in enumerate(appids, 1):
    r = S3.get("https://store.steampowered.com/api/appdetails",
               params={"appids": aid, "l": "english", "cc": "us"}, timeout=30)
    if r.status_code != 200:
        print(f"  requisicao #{i} (app {aid}) -> HTTP {r.status_code} "
              f"apos {time.time()-t0:.1f}s | Retry-After="
              f"{r.headers.get('Retry-After')!r}")
        break
    ok += 1
    time.sleep(0.35)
else:
    print(f"  {ok}/{len(appids)} respostas 200 em {time.time()-t0:.1f}s a ~0.35s/req "
          f"-> sem 429 nesta rajada")

path = os.path.join(OUT, "_funnel_counts.json")
data = json.load(open(path, encoding="utf-8")) if os.path.exists(path) else {}
data.setdefault("per_tag", {}).update({k: v for k, v in extra.items() if v is not None})
data["intersections"] = inter
with open(path, "w", encoding="utf-8") as fh:
    json.dump(data, fh, ensure_ascii=False, indent=2)
print(f"\n[DONE] funil atualizado em {path}")
