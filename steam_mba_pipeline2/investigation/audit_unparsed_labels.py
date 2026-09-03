"""Audita os rotulos de requisitos que o parser nao soube mapear.

Se um rotulo aparece com frequencia, e sinal de sinonimo faltando no
LABEL_SYNONYMS. Se aparece uma vez e e realmente exotico, `unparsed_labels`
cumpriu seu papel: preservou em vez de descartar em silencio.
"""
from __future__ import annotations

import json
import os
import sys
from collections import Counter

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "src"))

from steamreq.config import Filters, Settings  # noqa: E402
from steamreq.export import build_all_records  # noqa: E402

settings, filters = Settings.load(), Filters.load()
records = build_all_records(settings, filters)

cont: Counter[str] = Counter()
exemplos: dict[str, tuple[int, str, str]] = {}
for r in records:
    pcr = r.get("pc_requirements") or {}
    for bloco in ("minimum", "recommended"):
        for u in (pcr.get(bloco) or {}).get("unparsed_labels") or []:
            label = u["label"]
            cont[label] += 1
            exemplos.setdefault(label, (r["app_id"], bloco,
                                        str(u.get("value"))[:60]))

print(f"registros analisados: {len(records)}")
print(f"rotulos distintos nao mapeados: {len(cont)}\n")
print(f"{'ocorr':>6}  {'rotulo':<34} exemplo (app, bloco, valor)")
print("-" * 78)
for label, n in cont.most_common():
    aid, bloco, val = exemplos[label]
    print(f"{n:>6}  {label[:34]:<34} app={aid} {bloco} -> {val!r}")

# Cobertura efetiva dos campos, agora em escala.
print("\n" + "=" * 78)
print("COBERTURA DOS CAMPOS EM ESCALA")
print("=" * 78)
inc = [r for r in records if r.get("included_initially")]
com_rec = [r for r in inc if r.get("has_recommended_requirements")]
print(f"incluidos: {len(inc)} | com requisitos recomendados: {len(com_rec)} "
      f"({100*len(com_rec)/len(inc):.1f}%)")
for campo in ("os", "cpu", "ram", "gpu", "directx", "storage"):
    n = sum(1 for r in com_rec
            if ((r["pc_requirements"]["recommended"]) or {}).get(campo))
    print(f"  recommended_{campo:<9} {n:>4}/{len(com_rec)} "
          f"({100*n/len(com_rec):.1f}%)")

fmt: Counter[str] = Counter(r.get("requirements_markup_format") for r in inc)
print(f"\nformatos de markup encontrados: {dict(fmt)}")

anos: Counter[int] = Counter(r["release_year"] for r in inc if r.get("release_year"))
print(f"\ndistribuicao por ano de lancamento ({len(anos)} anos distintos):")
for ano in sorted(anos):
    print(f"  {ano}  {'#' * min(60, anos[ano])} {anos[ano]}")
