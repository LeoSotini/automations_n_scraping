"""Repara samples/_funnel_counts.json.

Motivo: probe_round4.py sofreu um TypeError de formatacao ANTES de chegar a
gravacao do arquivo, e probe_round5.py depois reescreveu o arquivo apenas com
as suas proprias medicoes. As contagens da rodada 4 existiam somente no output
do console.

Os valores abaixo sao transcritos literalmente daquele output — sao medicoes
reais de 2026-08-30, nao estimativas. A proveniencia fica registrada no proprio
arquivo, no campo `provenance`.
"""
from __future__ import annotations

import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(HERE, "samples", "_funnel_counts.json")

ROUND4_PER_TAG = {
    # tags positivas fortes (item 2.6)
    "3D": 51760, "First-Person": 30590, "Third Person": 18510,
    "Third-Person Shooter": 4617, "FPS": 10988, "3D Platformer": 9875,
    "3D Fighter": 2165, "Immersive Sim": 9954, "Walking Simulator": 7993,
    "Automobile Sim": 2211, "Flight": 3002, "Space Sim": 2083,
    "Looter Shooter": 1547, "Hero Shooter": 1339, "Arena Shooter": 4196,
    "Rail Shooter": 0,
    # tags positivas secundarias (item 2.7)
    "Open World": 13449, "Realistic": 14814, "Cinematic": 5740,
    "Action-Adventure": 25180, "Action RPG": 10695, "Souls-like": 3233,
    "Survival Horror": 8358, "Driving": 3718,
    # negativas
    "2D": 63469, "2D Platformer": 14965,
}

ROUND4_INTERSECTIONS = {
    "3D": 51760,
    "3D AND Indie": 25036,
    "3D AND 2D (conflito)": 1915,
    "3D AND 2.5D": 1071,
    "3D AND Pixel Graphics": 2389,
    "3D NOT Indie (untags)": 26724,
    "3D NOT Indie NOT 2D": 25575,
}

data = {}
if os.path.exists(PATH):
    with open(PATH, encoding="utf-8") as fh:
        data = json.load(fh)

per_tag = dict(ROUND4_PER_TAG)
per_tag.update(data.get("per_tag") or {})       # preserva o que a rodada 5 mediu

merged = {
    "provenance": {
        "measured_at": "2026-08-30",
        "method": "store.steampowered.com/search/results/ com "
                  "category1=998&os=win, total_count por tag",
        "note": "as contagens da rodada 4 foram restauradas a partir do output "
                "do console (o script falhou antes de gravar); as da rodada 5 "
                "vieram do proprio arquivo",
    },
    "baseline_windows_games": data.get("baseline_windows_games", 172017),
    "per_tag": dict(sorted(per_tag.items(), key=lambda kv: -kv[1])),
    "intersections": {**ROUND4_INTERSECTIONS, **(data.get("intersections") or {})},
}

with open(PATH, "w", encoding="utf-8") as fh:
    json.dump(merged, fh, ensure_ascii=False, indent=2)

print(f"[OK] {PATH}")
print(f"     baseline={merged['baseline_windows_games']:,}")
print(f"     tags com contagem: {len(merged['per_tag'])}")
print(f"     intersecoes: {len(merged['intersections'])}")
