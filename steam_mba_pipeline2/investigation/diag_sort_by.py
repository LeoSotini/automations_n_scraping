"""sort_by=Released_DESC altera o total_count do /search/results/?

Na FASE 1, "3D NOT Indie" media 26.724 (query SEM sort_by). O discovery, que
adiciona sort_by=Released_DESC, reportou 15.744 para o mesmo recorte.

Se sort_by reduz a populacao, e um vies silencioso: jogos sem data de
lancamento utilizavel sairiam do universo — e o corte de 2005 e justamente um
criterio de data. Precisa ser confirmado antes da FASE 5.
"""
from __future__ import annotations

import time

import requests

UA = "Mozilla/5.0 (academic-research)"
S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})
B = "https://store.steampowered.com/search/results/"


def count(**extra) -> int | None:
    p = {"query": "", "start": 0, "count": 1, "infinite": 1, "json": 1,
         "os": "win", "category1": 998}
    p.update(extra)
    for attempt in range(5):
        r = S.get(B, params=p, timeout=30)
        if r.status_code == 429:
            time.sleep(5 * (attempt + 1))
            continue
        time.sleep(4.0)
        try:
            return r.json().get("total_count")
        except Exception:  # noqa: BLE001
            return None
    return None


CASES = [
    ("3D, sem sort_by",                dict(tags=4191)),
    ("3D, sort_by=Released_DESC",      dict(tags=4191, sort_by="Released_DESC")),
    ("3D, sort_by=Relevance",          dict(tags=4191, sort_by="Relevance")),
    ("3D, sort_by=Reviews_DESC",       dict(tags=4191, sort_by="Reviews_DESC")),
    ("3D NOT Indie, sem sort_by",      dict(tags=4191, untags=492)),
    ("3D NOT Indie, Released_DESC",    dict(tags=4191, untags=492,
                                           sort_by="Released_DESC")),
    ("3D NOT Indie, Reviews_DESC",     dict(tags=4191, untags=492,
                                           sort_by="Reviews_DESC")),
]

print("=" * 78)
print("EFEITO DE sort_by NO total_count")
print("=" * 78)
res = {}
for label, params in CASES:
    c = count(**params)
    res[label] = c
    print(f"  {label:34s} -> {c:>8,}" if c is not None
          else f"  {label:34s} -> FALHOU")

print()
base = res.get("3D NOT Indie, sem sort_by")
srt = res.get("3D NOT Indie, Released_DESC")
if base and srt:
    print(f"  sem sort_by      = {base:,}")
    print(f"  Released_DESC    = {srt:,}")
    delta = base - srt
    print(f"  diferenca        = {delta:,} ({100*delta/base:.1f}% da populacao)")
    if abs(delta) > base * 0.02:
        print("\n  CONFIRMADO: sort_by=Released_DESC REDUZ a populacao.")
        print("  Provavel causa: itens sem data de lancamento ordenavel sao")
        print("  omitidos. Isso constitui vies silencioso e o parametro deve")
        print("  ser removido do discovery.")
    else:
        print("\n  sort_by NAO altera materialmente o total_count.")

print()
print("=" * 78)
print("A paginacao profunda funciona igual sem sort_by?")
print("=" * 78)
for start in (0, 5000, 15000):
    c = count(tags=4191, untags=492, start=start)
    print(f"  start={start:>6} -> total_count={c}")
