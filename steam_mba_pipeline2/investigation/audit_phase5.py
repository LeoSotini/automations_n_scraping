"""Auditoria do dataset completo da FASE 5.

Investiga as divergencias entre o piloto (310 jogos) e a coleta completa
(46.909), e verifica se o dataset esta analiticamente utilizavel.

Tudo offline: nenhuma requisicao a Steam.
"""
from __future__ import annotations

import os
import sys
from collections import Counter

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "src"))

from steamreq.config import Filters, Settings  # noqa: E402
from steamreq.export import build_all_records  # noqa: E402

settings, filters = Settings.load(), Filters.load()
print("carregando registros de data/raw ...")
records = build_all_records(settings, filters)
inc = [r for r in records if r.get("included_initially")]
exc = [r for r in records if not r.get("included_initially")]
print(f"total={len(records):,} | incluidos={len(inc):,} | excluidos={len(exc):,}\n")


def h(titulo: str) -> None:
    print("=" * 78)
    print(titulo)
    print("=" * 78)


# --- 1. ANOMALIA: 17.368 UNRELEASED (37% da coleta) -----------------------
h("1) UNRELEASED — 17.368 no total (37%). No piloto de 310 era ZERO.")
unrel = [r for r in exc if r.get("exclusion_reason") == "UNRELEASED"]
print(f"  registros UNRELEASED: {len(unrel):,} "
      f"({100*len(unrel)/len(records):.1f}% da coleta)")
print(f"  todos tem coming_soon=True? "
      f"{all(r.get('coming_soon') for r in unrel)}")
datas = Counter(r.get("release_date_raw") or "(vazio)" for r in unrel)
print(f"\n  release_date_raw mais frequentes entre os UNRELEASED:")
for d, n in datas.most_common(14):
    print(f"    {n:>6,}  {d!r}")
com_req = sum(1 for r in unrel if r.get("has_recommended_requirements"))
print(f"\n  desses, quantos JA publicam requisitos recomendados: "
      f"{com_req:,} ({100*com_req/len(unrel):.1f}%)")
print("  -> requisitos de jogos nao lancados sao PROVISORIOS (item 2.4 do")
print("     projeto ja previa isso). A exclusao esta correta.")
print(f"\n  CUSTO: ~{2*len(unrel):,} requisicoes gastas em jogos que seriam")
print("  descartados. Ver se o discovery pode evita-los na fonte.")

# --- 2. BEFORE_2005 = apenas 14 ------------------------------------------
h("2) BEFORE_2005 — apenas 14. Plausivel ou o filtro falhou?")
antigos = [r for r in exc if r.get("exclusion_reason") == "BEFORE_2005"]
for r in sorted(antigos, key=lambda x: x.get("release_date") or ""):
    print(f"    {r.get('release_date')}  {str(r.get('name'))[:46]}")
lancados = [r for r in records if not r.get("coming_soon")
            and r.get("release_year")]
anos = Counter(r["release_year"] for r in lancados)
print(f"\n  jogos lancados por ano (os mais antigos):")
for ano in sorted(anos)[:12]:
    print(f"    {ano}  {anos[ano]:>6,}")
print("  -> a Steam tinha pouquissimos jogos antes de 2005; o volume so cresce")
print("     a partir de 2012. 14 exclusoes e coerente com o catalogo real.")

# --- 3. 21% sem requisitos recomendados (piloto: 4,6%) -------------------
h("3) SEM requisitos recomendados — 21,0% agora, 4,6% no piloto. Por que?")
sem_rec = [r for r in inc if not r.get("has_recommended_requirements")]
print(f"  incluidos sem recomendados: {len(sem_rec):,} de {len(inc):,} "
      f"({100*len(sem_rec)/len(inc):.1f}%)")
por_ano_sem: Counter[int] = Counter()
por_ano_tot: Counter[int] = Counter()
for r in inc:
    a = r.get("release_year")
    if a:
        por_ano_tot[a] += 1
        if not r.get("has_recommended_requirements"):
            por_ano_sem[a] += 1
print(f"\n  taxa de ausencia por ano de lancamento:")
for a in sorted(por_ano_tot):
    if por_ano_tot[a] >= 30:
        pct = 100 * por_ano_sem[a] / por_ano_tot[a]
        print(f"    {a}  {por_ano_sem[a]:>5,}/{por_ano_tot[a]:>6,}  "
              f"{pct:>5.1f}%  {'#' * int(pct/2)}")
rev_com = [r.get("review_count") or 0 for r in inc
           if r.get("has_recommended_requirements")]
rev_sem = [r.get("review_count") or 0 for r in sem_rec]
med = lambda xs: sorted(xs)[len(xs)//2] if xs else 0  # noqa: E731
print(f"\n  mediana de reviews COM recomendados : {med(rev_com):,}")
print(f"  mediana de reviews SEM recomendados : {med(rev_sem):,}")
print("  -> o piloto pegou as primeiras paginas de 3 tags (jogos mais")
print("     proeminentes). A populacao completa tem muito mais titulos")
print("     obscuros, que frequentemente nao publicam recomendados.")

# --- 4. UNPARSED_LABELS em escala: 2.017 avisos --------------------------
h("4) UNPARSED_LABELS — 2.017 avisos. Ha sinonimo faltando?")
cont: Counter[str] = Counter()
exemplo: dict[str, tuple[int, str]] = {}
for r in records:
    pcr = r.get("pc_requirements") or {}
    for bloco in ("minimum", "recommended"):
        for u in (pcr.get(bloco) or {}).get("unparsed_labels") or []:
            cont[u["label"]] += 1
            exemplo.setdefault(u["label"], (r["app_id"], str(u.get("value"))[:50]))
print(f"  rotulos distintos nao mapeados: {len(cont):,}")
print(f"  ocorrencias totais: {sum(cont.values()):,}\n")
print(f"  {'ocorr':>6}  {'rotulo':<38} exemplo")
print("  " + "-" * 74)
for label, n in cont.most_common(30):
    aid, val = exemplo[label]
    print(f"  {n:>6,}  {label[:38]:<38} app={aid} {val!r}")

# --- 5. Avisos residuais -------------------------------------------------
h("5) INDIE_PASSED_SOURCE_FILTER (31) e INCLUDED_WITHOUT_DATE (15)")
indies = [r for r in inc if r.get("is_indie")]
print(f"  incluidos com is_indie=True: {len(indies)}")
for r in indies[:12]:
    print(f"    app={r['app_id']:<9} {str(r.get('name'))[:34]:<36} "
          f"via={r.get('discovered_via_tags') or '?'} "
          f"tags2d={r.get('has_2d_tags')}")
sem_data = [r for r in inc if not r.get("release_date")]
print(f"\n  incluidos sem data utilizavel: {len(sem_data)}")
for r in sem_data[:15]:
    print(f"    app={r['app_id']:<9} raw={str(r.get('release_date_raw'))[:24]!r:<26} "
          f"coming_soon={r.get('coming_soon')}")

# --- 6. O dataset e analiticamente utilizavel? --------------------------
h("6) UTILIDADE ANALITICA — cobertura dos campos recomendados")
com_rec = [r for r in inc if r.get("has_recommended_requirements")]
print(f"  incluidos com recomendados: {len(com_rec):,}")
for campo in ("os", "cpu", "ram", "gpu", "directx", "storage"):
    n = sum(1 for r in com_rec
            if (r["pc_requirements"]["recommended"] or {}).get(campo))
    print(f"    recommended_{campo:<9} {n:>7,}/{len(com_rec):,} "
          f"({100*n/len(com_rec):.1f}%)")

print(f"\n  registros com RAM e ano (nucleo da analise longitudinal):")
uteis = [r for r in com_rec
         if (r["pc_requirements"]["recommended"] or {}).get("ram")
         and r.get("release_year")]
print(f"    {len(uteis):,} registros")
anos_u = Counter(r["release_year"] for r in uteis)
for a in sorted(anos_u):
    if a >= 2005:
        print(f"    {a}  {anos_u[a]:>5,}  {'#' * min(58, anos_u[a]//12)}")

fmt = Counter(r.get("requirements_markup_format") for r in inc)
print(f"\n  formatos de markup: {dict(fmt)}")
basis = Counter(r.get("inclusion_basis") for r in inc)
print(f"  base de inclusao: {dict(basis)}")
src = Counter(r.get("tags_source") for r in inc)
print(f"  fonte das tags: {dict(src)}")
