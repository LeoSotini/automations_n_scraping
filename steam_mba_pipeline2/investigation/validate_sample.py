"""FASE 4 - Validacao manual assistida.

Confronta os campos RECOMENDADOS coletados com o texto bruto que a Steam
publica, item por item. Imprime lado a lado para conferencia visual, e checa
automaticamente que cada valor estruturado aparece literalmente no raw.
"""
from __future__ import annotations

import json
import os
import re
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATASET = os.path.join(HERE, "data", "processed", "dataset.json")

with open(DATASET, encoding="utf-8") as fh:
    records = json.load(fh)


def visible(html: str | None) -> str:
    if not html:
        return ""
    t = re.sub(r"<br\s*/?>", " ", html, flags=re.IGNORECASE)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t).strip()


FIELDS = ["os", "cpu", "ram", "gpu", "directx", "storage"]
problemas: list[str] = []

print("=" * 78)
print("FASE 4 - CONFERENCIA DOS REQUISITOS RECOMENDADOS")
print("=" * 78)
print(f"registros no dataset: {len(records)}\n")

com_rec = [r for r in records if r.get("has_recommended_requirements")]
sem_rec = [r for r in records if not r.get("has_recommended_requirements")]

for r in sorted(com_rec, key=lambda x: x.get("release_year") or 0):
    print("-" * 78)
    print(f"{r['name']}  (app {r['app_id']}, {r['release_date']}, "
          f"formato {r['requirements_markup_format']})")
    print(f"  URL: {r['steam_url']}")
    print(f"  RAW RECOMENDADO (texto visivel na Steam):")
    print(f"    {visible(r.get('recommended_raw'))[:330]}")
    print("  CAMPOS EXTRAIDOS:")
    for f in FIELDS:
        val = r.get(f"recommended_{f}")
        marca = " " if val else "-"
        print(f"    {marca} recommended_{f:<9} = {val!r}")
        # Checagem automatica: o valor deve existir literalmente no raw.
        if val and val not in visible(r.get("recommended_raw")):
            problemas.append(f"{r['name']}: recommended_{f}={val!r} nao aparece "
                             "no raw")
    # Confronto minimo vs recomendado: nao podem ser iguais por copia.
    iguais = [f for f in FIELDS
              if r.get(f"minimum_{f}") is not None
              and r.get(f"minimum_{f}") == r.get(f"recommended_{f}")]
    print(f"    campos identicos entre minimo e recomendado: {iguais or 'nenhum'}")
    if r.get("recommended_unparsed_labels"):
        print(f"    ATENCAO rotulos nao mapeados: "
              f"{[u['label'] for u in r['recommended_unparsed_labels']]}")
    if r.get("recommended_os_legacy_flag"):
        print("    os_legacy_flag=True (rotulo 'OS *' na Steam)")
    print(f"    tags ({r.get('tags_source')}): "
          f"{(r.get('steam_tags') or [])[:8]}")
    print(f"    inclusion_basis={r.get('inclusion_basis')} "
          f"| matched={r.get('matched_positive_tags')}")

print("-" * 78)
print(f"\nSEM requisitos recomendados na Steam ({len(sem_rec)}):")
for r in sem_rec:
    tem_min = bool(r.get("minimum_raw"))
    print(f"  {r['name']:<42} app={r['app_id']:<9} "
          f"has_recommended={r['has_recommended_requirements']} "
          f"minimos_presentes={tem_min}")
    if r.get("recommended_raw") is not None:
        problemas.append(f"{r['name']}: has_recommended=False mas raw presente")
    for f in FIELDS:
        if r.get(f"recommended_{f}") is not None:
            problemas.append(f"{r['name']}: recommended_{f} preenchido sem "
                             "requisitos recomendados na Steam")

print("\n" + "=" * 78)
print("CHECAGENS AUTOMATICAS")
print("=" * 78)
if problemas:
    for p in problemas:
        print(f"  PROBLEMA: {p}")
else:
    print("  nenhum problema: todo valor estruturado aparece literalmente no")
    print("  raw, e nenhum campo recomendado foi preenchido sem fonte real")

# Cobertura de extracao por campo, apenas entre quem tem recomendados.
print("\nCOBERTURA DOS CAMPOS RECOMENDADOS (entre os "
      f"{len(com_rec)} que os possuem):")
for f in FIELDS:
    n = sum(1 for r in com_rec if r.get(f"recommended_{f}"))
    print(f"  recommended_{f:<9} {n:>2}/{len(com_rec)} "
          f"({100*n/len(com_rec):.0f}%)")

sys.exit(1 if problemas else 0)
