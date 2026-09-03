"""Investiga as lacunas de ~5% nos campos recomendados.

Objetivo: separar ausencia REAL na Steam de falha do parser. A distincao
importa porque a primeira e um dado, e a segunda e um bug.
"""
from __future__ import annotations

import os
import re
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "src"))

from steamreq.config import Filters, Settings  # noqa: E402
from steamreq.export import build_all_records  # noqa: E402

settings, filters = Settings.load(), Filters.load()
records = build_all_records(settings, filters)
inc = [r for r in records if r.get("included_initially")
       and r.get("has_recommended_requirements")]


def visible(html: str | None) -> str:
    if not html:
        return ""
    t = re.sub(r"<br\s*/?>", " ", html, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t)).strip()


for campo, palavras in (("cpu", ("processor", "cpu")),
                        ("ram", ("memory", "ram")),
                        ("gpu", ("graphics", "video")),
                        ("storage", ("storage", "hard", "space", "disk")),
                        ("os", ("os", "operating"))):
    faltando = [r for r in inc
                if not (r["pc_requirements"]["recommended"] or {}).get(campo)]
    if not faltando:
        continue
    print("=" * 78)
    print(f"CAMPO recommended_{campo}: {len(faltando)} ausentes de {len(inc)}")
    print("=" * 78)
    real, suspeito = 0, []
    for r in faltando:
        raw = visible(r["pc_requirements"]["recommended"].get("raw")).lower()
        # Se o raw menciona a palavra-chave, o parser DEVERIA ter extraido.
        if any(p in raw for p in palavras):
            suspeito.append(r)
        else:
            real += 1
    print(f"  ausencia REAL na Steam (raw nem menciona o conceito): {real}")
    print(f"  SUSPEITO de falha de parser: {len(suspeito)}")
    for r in suspeito[:6]:
        print(f"\n    app {r['app_id']} — {r['name']}")
        print(f"    raw: {visible(r['pc_requirements']['recommended']['raw'])[:260]}")
        print(f"    unparsed: {[u['label'] for u in r['pc_requirements']['recommended']['unparsed_labels']]}")
    print()
