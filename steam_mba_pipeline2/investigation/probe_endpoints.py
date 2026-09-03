"""FASE 1 - Sondagem empirica de fontes publicas da Steam.

Objetivo: verificar QUAIS endpoints estao funcionando HOJE, quais campos
entregam, e como se comportam em casos de borda (age gate, jogo removido,
jogo sem requisitos recomendados, jogo antigo, F2P).

Nao coleta em escala. Apenas ~20 requisicoes com pausas.
"""
from __future__ import annotations

import json
import os
import time
from typing import Any

import requests

OUT = os.path.join(os.path.dirname(__file__), "samples")
os.makedirs(OUT, exist_ok=True)

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36 "
    "(academic-research; MBA DSA TCC; contact: leonardo.de-lima@siemens-energy.com)"
)

S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})

PAUSE = 1.5
report: list[dict[str, Any]] = []


def probe(label: str, url: str, *, params=None, cookies=None, save: str | None = None,
          expect_json: bool = True, allow_redirects: bool = True) -> dict[str, Any]:
    entry: dict[str, Any] = {"label": label, "url": url, "params": params}
    t0 = time.time()
    try:
        r = S.get(url, params=params, cookies=cookies, timeout=30,
                  allow_redirects=allow_redirects)
        entry["status"] = r.status_code
        entry["elapsed_s"] = round(time.time() - t0, 2)
        entry["final_url"] = r.url
        entry["content_type"] = r.headers.get("Content-Type")
        entry["bytes"] = len(r.content)
        entry["rate_headers"] = {
            k: v for k, v in r.headers.items()
            if "rate" in k.lower() or "retry" in k.lower()
        }
        body = None
        if expect_json:
            try:
                body = r.json()
                entry["json_ok"] = True
            except Exception as e:  # noqa: BLE001
                entry["json_ok"] = False
                entry["json_error"] = str(e)[:200]
        if save:
            path = os.path.join(OUT, save)
            if body is not None:
                with open(path, "w", encoding="utf-8") as fh:
                    json.dump(body, fh, ensure_ascii=False, indent=2)
            else:
                with open(path, "w", encoding="utf-8", errors="replace") as fh:
                    fh.write(r.text)
            entry["saved"] = save
        entry["body"] = body
    except Exception as e:  # noqa: BLE001
        entry["error"] = f"{type(e).__name__}: {e}"
    report.append(entry)
    short = {k: v for k, v in entry.items() if k != "body"}
    print(f"[PROBE] {label}\n        {json.dumps(short, ensure_ascii=False)[:600]}\n")
    time.sleep(PAUSE)
    return entry


# --- Casos de teste heterogeneos -------------------------------------------
CASES = {
    "csgo_f2p": 730,            # F2P, muitas reviews
    "gtav_agegate": 271590,     # age gate (mature)
    "hl2_2004": 220,            # ANTES de 2005 (teste de corte de data)
    "cyberpunk_2020": 1091500,  # AAA recente
    "hades_indie_2d": 1145360,  # Indie / 2.5D (teste de tags negativas)
    "terraria_2d": 105600,      # 2D puro
    "portal2_2011": 620,        # antigo-medio
    "eldenring_2022": 1245620,  # AAA, souls-like
    "dlc_example": 2138330,     # candidato a DLC (nao-game)
    "removed_or_bogus": 999999999,  # app_id inexistente
}

print("=" * 78)
print("1) DISCOVERY: ISteamApps/GetAppList/v2 (sem API key)")
print("=" * 78)
applist = probe(
    "GetAppList v2",
    "https://api.steampowered.com/ISteamApps/GetAppList/v2/",
    save="applist_v2.json",
)
if applist.get("body"):
    apps = applist["body"].get("applist", {}).get("apps", [])
    print(f"        -> total de entradas na applist: {len(apps):,}")
    print(f"        -> exemplo: {apps[:3]}")
    print(f"        -> campos disponiveis: {sorted(apps[0].keys()) if apps else None}")

print("=" * 78)
print("2) DISCOVERY alternativo: IStoreService/GetAppList/v1 (paginado, precisa key?)")
print("=" * 78)
probe(
    "IStoreService GetAppList v1 (sem key)",
    "https://api.steampowered.com/IStoreService/GetAppList/v1/",
    params={"include_games": "true", "max_results": 10},
    save="istoreservice_nokey.json",
)

print("=" * 78)
print("3) METADADOS + REQUISITOS: store appdetails")
print("=" * 78)
for label, appid in CASES.items():
    probe(
        f"appdetails {label} ({appid})",
        "https://store.steampowered.com/api/appdetails",
        params={"appids": appid, "l": "english", "cc": "us"},
        save=f"appdetails_{label}_{appid}.json",
    )

print("=" * 78)
print("4) REVIEW COUNT: appreviews (num_per_page=0 -> so o sumario)")
print("=" * 78)
for label, appid in [("csgo", 730), ("cyberpunk", 1091500), ("bogus", 999999999)]:
    probe(
        f"appreviews {label}",
        f"https://store.steampowered.com/appreviews/{appid}",
        params={"json": 1, "num_per_page": 0, "language": "all",
                "purchase_type": "all"},
        save=f"appreviews_{label}_{appid}.json",
    )

print("=" * 78)
print("5) TAGS: appdetails NAO retorna steam_tags. Testando alternativas")
print("=" * 78)
# 5a) SteamSpy (terceiro, mas expoe tags agregadas + owners)
probe(
    "steamspy appdetails 1091500",
    "https://steamspy.com/api.php",
    params={"request": "appdetails", "appid": 1091500},
    save="steamspy_1091500.json",
)
# 5b) HTML da store: bloco de tags app_tag / InitAppTagModal
probe(
    "store HTML 1091500 (tags via HTML)",
    "https://store.steampowered.com/app/1091500/",
    cookies={"birthtime": "283993201", "lastagecheckage": "1-January-1979",
             "wants_mature_content": "1", "Steam_Language": "english"},
    save="store_html_1091500.html",
    expect_json=False,
)
# 5c) age gate SEM cookies
probe(
    "store HTML 271590 SEM cookie de idade",
    "https://store.steampowered.com/app/271590/",
    save="store_html_271590_noagecookie.html",
    expect_json=False,
)

print("=" * 78)
print("6) SEARCH paginado (discovery filtrado por tag direto na Steam)")
print("=" * 78)
probe(
    "search results infinite (tag 3D=4166, Windows, ordenado)",
    "https://store.steampowered.com/search/results/",
    params={"query": "", "start": 0, "count": 50, "infinite": 1,
            "category1": 998, "tags": 4166, "os": "win",
            "supportedlang": "english", "json": 1},
    save="search_tag3d.json",
)

with open(os.path.join(OUT, "_probe_report.json"), "w", encoding="utf-8") as fh:
    json.dump([{k: v for k, v in e.items() if k != "body"} for e in report],
              fh, ensure_ascii=False, indent=2)
print(f"\n[DONE] {len(report)} sondagens. Amostras em {OUT}")
