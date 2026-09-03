"""FASE 1b - Variantes de discovery + analise do conteudo das amostras."""
from __future__ import annotations

import glob
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

print("=" * 78)
print("A) VARIANTES DE DISCOVERY")
print("=" * 78)
variants = [
    ("v2 sem barra", "https://api.steampowered.com/ISteamApps/GetAppList/v2", None),
    ("v0002", "https://api.steampowered.com/ISteamApps/GetAppList/v0002/",
     {"format": "json"}),
    ("v1", "https://api.steampowered.com/ISteamApps/GetAppList/v1/", None),
    ("store actions GetAppList game",
     "https://store.steampowered.com/actions/GetAppList/",
     {"appType": "game"}),
    ("api.steampowered ISteamApps/GetAppList sem versao",
     "https://api.steampowered.com/ISteamApps/GetAppList/", None),
]
for label, url, params in variants:
    try:
        r = S.get(url, params=params, timeout=30)
        ok = None
        n = None
        try:
            j = r.json()
            ok = True
            apps = j.get("applist", {}).get("apps")
            if isinstance(apps, dict):
                apps = apps.get("app")
            n = len(apps) if isinstance(apps, list) else None
            if n:
                with open(os.path.join(OUT, "applist_WORKING.json"), "w",
                          encoding="utf-8") as fh:
                    json.dump(j, fh)
        except Exception:
            ok = False
        print(f"  {label:52s} -> HTTP {r.status_code} json={ok} apps={n} "
              f"bytes={len(r.content)}")
    except Exception as e:  # noqa: BLE001
        print(f"  {label:52s} -> ERRO {type(e).__name__}: {e}")
    time.sleep(1.2)

print()
print("=" * 78)
print("B) ANALISE DAS AMOSTRAS appdetails")
print("=" * 78)
for path in sorted(glob.glob(os.path.join(OUT, "appdetails_*.json"))):
    name = os.path.basename(path)
    with open(path, encoding="utf-8") as fh:
        j = json.load(fh)
    key = next(iter(j))
    node = j[key]
    if not node.get("success"):
        print(f"\n  {name}\n    success=False  (payload={json.dumps(j)[:120]})")
        continue
    d = node["data"]
    pcr = d.get("pc_requirements")
    pcr_kind = type(pcr).__name__
    has_min = isinstance(pcr, dict) and bool(pcr.get("minimum"))
    has_rec = isinstance(pcr, dict) and bool(pcr.get("recommended"))
    print(f"\n  {name}")
    print(f"    type={d.get('type')!r} name={d.get('name')!r}")
    print(f"    release={d.get('release_date')}")
    print(f"    platforms={d.get('platforms')}")
    print(f"    genres={[g['description'] for g in d.get('genres', [])]}")
    print(f"    categories={[c['description'] for c in d.get('categories', [])][:8]}")
    print(f"    devs={d.get('developers')} pubs={d.get('publishers')}")
    print(f"    required_age={d.get('required_age')!r} is_free={d.get('is_free')}")
    print(f"    n_languages_field_len={len(d.get('supported_languages') or '')}")
    print(f"    pc_requirements: kind={pcr_kind} min={has_min} rec={has_rec}")
    print(f"    TEM campo de tags? {'tags' in d or 'steam_tags' in d}")
    print(f"    top-level keys={sorted(d.keys())}")
    if has_rec:
        raw = pcr["recommended"]
        print(f"    RECOMMENDED RAW (400 chars):\n      {raw[:400]}")
    elif has_min:
        print(f"    MINIMUM RAW (300 chars):\n      {pcr['minimum'][:300]}")

print()
print("=" * 78)
print("C) appreviews - sumario")
print("=" * 78)
for path in sorted(glob.glob(os.path.join(OUT, "appreviews_*.json"))):
    with open(path, encoding="utf-8") as fh:
        j = json.load(fh)
    print(f"  {os.path.basename(path)} -> success={j.get('success')} "
          f"query_summary={json.dumps(j.get('query_summary'))}")

print()
print("=" * 78)
print("D) TAGS no HTML da store (Cyberpunk 1091500)")
print("=" * 78)
html_path = os.path.join(OUT, "store_html_1091500.html")
with open(html_path, encoding="utf-8", errors="replace") as fh:
    html = fh.read()
m = re.search(r"InitAppTagModal\(\s*\d+\s*,\s*(\[.*?\])\s*,", html, re.S)
if m:
    tags = json.loads(m.group(1))
    print(f"  InitAppTagModal ENCONTRADO -> {len(tags)} tags")
    print(f"  primeiras: {[(t['tagid'], t['name']) for t in tags[:20]]}")
else:
    print("  InitAppTagModal NAO encontrado")
vis = re.findall(r'app_tag[^>]*>\s*([^<]+?)\s*<', html)
print(f"  tags visiveis no HTML (classe app_tag): {[v for v in vis][:20]}")
print(f"  tem 'agegate'/'agecheck' no HTML? "
      f"{'agecheck' in html.lower() or 'agegate' in html.lower()}")

print()
print("=" * 78)
print("E) AGE GATE sem cookie (GTA V 271590)")
print("=" * 78)
p = os.path.join(OUT, "store_html_271590_noagecookie.html")
with open(p, encoding="utf-8", errors="replace") as fh:
    h2 = fh.read()
print(f"  bytes={len(h2)}")
print(f"  contem 'agecheck'? {'agecheck' in h2.lower()}")
print(f"  contem 'InitAppTagModal'? {'InitAppTagModal' in h2}")
print(f"  contem 'game_area_sys_req'? {'game_area_sys_req' in h2}")
print(f"  <title>: {re.search(r'<title>(.*?)</title>', h2, re.S).group(1).strip()[:120]}")

print()
print("=" * 78)
print("F) SEARCH JSON - estrutura")
print("=" * 78)
with open(os.path.join(OUT, "search_tag3d.json"), encoding="utf-8") as fh:
    sj = json.load(fh)
print(f"  keys={sorted(sj.keys())}")
print(f"  total_count={sj.get('total_count')}  results_html_len="
      f"{len(sj.get('results_html') or '')}")
ids = re.findall(r'data-ds-appid="(\d+)"', sj.get("results_html", ""))
print(f"  appids extraidos desta pagina: {len(ids)} -> {ids[:15]}")
