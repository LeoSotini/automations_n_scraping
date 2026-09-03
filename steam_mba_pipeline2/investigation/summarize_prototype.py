"""Resumo do resultado final do prototipo (FASES 3/4 + decisao D10)."""
import json
import os

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
P = os.path.join(HERE, "data", "processed")

with open(os.path.join(P, "ledger_exclusions.json"), encoding="utf-8") as fh:
    led = json.load(fh)

print("LEDGER DE EXCLUSOES (item 2.17) — nenhum descarte sem motivo:")
for r in led["records"]:
    tags = (r.get("steam_tags") or [])[:5]
    print(f"  {str(r.get('name'))[:44]:<46} {r['exclusion_reason']:<18} {tags}")

with open(os.path.join(P, "dataset.json"), encoding="utf-8") as fh:
    ds = json.load(fh)

print(f"\nDATASET FINAL ({len(ds)} registros incluidos):")
for r in sorted(ds, key=lambda x: x.get("release_year") or 0):
    print(f"  {r.get('release_year')}  {r['name'][:40]:<42} "
          f"rec={str(r['has_recommended_requirements']):<5} "
          f"basis={r.get('inclusion_basis')}")

com = [r for r in ds if r["has_recommended_requirements"]]
print(f"\nEvolucao de RAM/armazenamento recomendados (o dado central do estudo):")
for r in sorted(com, key=lambda x: x.get("release_year") or 0):
    print(f"  {r.get('release_year')}  {r['name'][:32]:<34} "
          f"RAM={str(r.get('recommended_ram'))[:14]:<16} "
          f"storage={str(r.get('recommended_storage'))[:26]}")
