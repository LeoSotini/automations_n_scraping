"""Diagnostico: por que InitAppTagModal desapareceu em 7 de 17 paginas?

Na FASE 1, /app/1091500/ retornou 267 KB COM as tags. No prototipo, a mesma URL
retornou 51 KB SEM as tags. Algo mudou entre as duas execucoes. Hipoteses a
testar:
  H1 - age gate renderizado NA PAGINA (nao por redirect)
  H2 - falta de cookies que a sondagem da FASE 1 enviava
  H3 - resposta degradada por ritmo de requisicao (anti-bot leve)
  H4 - variante de pagina por regiao/idioma
"""
from __future__ import annotations

import re
import sys
import time

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36 (academic-research)")

FAILED = [1091500, 1245620, 1145360, 292030, 413150]
WORKED = [271590, 1174180, 582010]

AGE_COOKIES = {"birthtime": "283993201", "lastagecheckage": "1-January-1979",
               "wants_mature_content": "1", "Steam_Language": "english"}


def probe(app_id: int, label: str, cookies=None, session=None):
    s = session or requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})
    r = s.get(f"https://store.steampowered.com/app/{app_id}/",
              cookies=cookies, timeout=30, allow_redirects=True)
    html = r.text
    low = html.lower()
    has_modal = "InitAppTagModal" in html
    title = re.search(r"<title>(.*?)</title>", html, re.S)
    print(f"  app {app_id:<9} {label:<26} HTTP {r.status_code} "
          f"{len(r.content)/1024:6.1f} KB modal={has_modal!s:<5} "
          f"agecheck={'agecheck' in low!s:<5} "
          f"sysreq={'game_area_sys_req' in html!s:<5}")
    print(f"      title={title.group(1).strip()[:70] if title else '?'!r}")
    if not has_modal:
        for marker in ("age_gate", "agegate", "mature_content", "app_tag",
                       "glance_tags", "error_box", "Cookies", "consent"):
            if marker.lower() in low:
                print(f"      marcador presente: {marker}")
    return has_modal, len(r.content), html


print("=" * 78)
print("H1/H2 — cookies de idade fazem diferenca?")
print("=" * 78)
for aid in FAILED[:3]:
    probe(aid, "SEM cookies")
    time.sleep(1.5)
    probe(aid, "COM cookies de idade", cookies=AGE_COOKIES)
    time.sleep(1.5)

print()
print("=" * 78)
print("H3 — sessao reutilizada vs nova a cada requisicao")
print("=" * 78)
shared = requests.Session()
for aid in FAILED[:3]:
    probe(aid, "sessao COMPARTILHADA", session=shared)
    time.sleep(1.5)

print()
print("=" * 78)
print("Controle: paginas que funcionaram no prototipo")
print("=" * 78)
for aid in WORKED:
    probe(aid, "controle SEM cookies")
    time.sleep(1.5)

print()
print("=" * 78)
print("Amostra do HTML de 51 KB (o que veio no lugar das tags?)")
print("=" * 78)
_, size, html = probe(1091500, "captura para inspecao")
body = re.sub(r"<script.*?</script>", "", html, flags=re.S)
text = re.sub(r"<[^>]+>", " ", body)
text = re.sub(r"\s+", " ", text).strip()
print(f"  primeiros 700 chars do texto visivel:\n  {text[:700]}")
out = "investigation/samples/diag_51kb_page.html"
with open(out, "w", encoding="utf-8", errors="replace") as fh:
    fh.write(html)
print(f"\n  HTML salvo em {out}")
