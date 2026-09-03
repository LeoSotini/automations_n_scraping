"""Estimativa do custo da FASE 5 em requisicoes e tempo.

Usa as contagens REAIS medidas na FASE 1 (samples/_funnel_counts.json) e as
configuracoes atuais de settings.yaml. Nao faz nenhuma requisicao.
"""
from __future__ import annotations

import json
import os
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "src"))

import yaml  # noqa: E402

from steamreq.config import Filters, Settings  # noqa: E402

settings = Settings.load()
filters = Filters.load()

with open(os.path.join(HERE, "investigation", "samples", "_funnel_counts.json"),
          encoding="utf-8") as fh:
    funnel = json.load(fh)
per_tag = funnel["per_tag"]

positive = filters.positive_strong + filters.positive_secondary
known = {t: per_tag[t] for t in positive if t in per_tag and per_tag[t]}
faltando = [t for t in positive if t not in known]

# Fracao Indie medida diretamente para a tag 3D: 25.036 de 51.760 = 48,4%.
# Usada como aproximacao para as demais (nao ha medicao por tag).
FRAC_INDIE = 25036 / 51760

soma_com_indie = sum(known.values())
soma_sem_indie = soma_com_indie * (1 - FRAC_INDIE)

page_size = settings.discovery["page_size"]
t_search = settings.rate_limits["search"]["min_interval_s"]
t_app = settings.rate_limits["appdetails"]["min_interval_s"]
t_html = settings.rate_limits["store_html"]["min_interval_s"]
sleep_app = settings.scrape.get("sleep_between_apps_s", 0.0)
fetch_tags = settings.scrape.get("fetch_tags_from_html", True)

print("=" * 78)
print("ESTIMATIVA DE CUSTO DA FASE 5")
print("=" * 78)
print(f"tags positivas com contagem medida : {len(known)}/{len(positive)}")
if faltando:
    print(f"  sem medicao (serao consultadas de todo modo): {faltando}")
print(f"soma das contagens por tag         : {soma_com_indie:,.0f} linhas")
print(f"  descontando Indie (~{100*FRAC_INDIE:.1f}%)      : {soma_sem_indie:,.0f} linhas")
print()

print("-" * 78)
print("ESTAGIO 1 — DISCOVERY")
print("-" * 78)
pag_princ = soma_sem_indie / page_size
pag_ledger = (soma_com_indie - soma_sem_indie) / page_size
h_princ = pag_princ * t_search / 3600
h_ledger = pag_ledger * t_search / 3600
print(f"  paginas do discovery principal   : {pag_princ:,.0f} "
      f"({t_search:.1f}s/req) -> {h_princ:.1f} h")
print(f"  paginas do ledger de Indie       : {pag_ledger:,.0f} -> {h_ledger:.1f} h")
print(f"  TOTAL discovery                  : {h_princ + h_ledger:.1f} h")
print()

print("-" * 78)
print("ESTAGIO 2 — SCRAPE")
print("-" * 78)
# Sobreposicao entre tags: um jogo aparece em varias. Cenarios de deduplicacao.
for label, frac in (("otimista (muita sobreposicao)", 0.35),
                    ("intermediario", 0.50),
                    ("pessimista (pouca sobreposicao)", 0.70)):
    n = soma_sem_indie * frac
    req_por_app = 2 if fetch_tags else 1
    t_req = t_app + (t_html if fetch_tags else 0)
    h = n * (sleep_app + t_req) / 3600
    print(f"  {label:32s} {n:>8,.0f} jogos | "
          f"{n*req_por_app:>9,.0f} reqs | {h:>6.1f} h ({h/24:>4.1f} dias)")
print()
print(f"  configuracao atual: sleep_between_apps_s={sleep_app:.1f}s + "
      f"{t_app:.1f}s (appdetails) + {t_html:.1f}s (html)")
print()

print("-" * 78)
print("COMPARACAO DE RITMOS (cenario intermediario)")
print("-" * 78)
n = soma_sem_indie * 0.50
print(f"  populacao assumida: {n:,.0f} jogos, 2 requisicoes cada\n")
print(f"  {'sleep entre jogos':<22} {'tempo total':>14} {'req/s efetivo':>16}")
for s in (0.0, 1.0, 2.0, 3.0, 5.0, 8.0):
    t_req = t_app + t_html
    total_s = n * (s + t_req)
    rps = (n * 2) / total_s
    marca = "  <- configurado" if abs(s - sleep_app) < 0.01 else ""
    print(f"  {s:>5.1f}s{'':<16} {total_s/3600:>10.1f} h "
          f"({total_s/86400:>4.1f} d) {rps:>10.2f}{marca}")
print()
print("  Referencia empirica da FASE 1: /api/appdetails respondeu 30/30 (HTTP")
print("  200) a 0,35s/req, isto e ~2,9 req/s, sem nenhum 429. O endpoint que")
print("  efetivamente limitou foi o /search/results/ (429 na 17a requisicao),")
print("  usado apenas no discovery — e ja configurado a 4,0s/req.")
