"""Fixtures pequenas, extraidas dos payloads REAIS coletados na FASE 1.

Origem: investigation/samples/appdetails_*.json. Mantidas curtas de proposito
(item 13: "utilize exemplos reais... mas mantenha fixtures pequenas").
"""
from __future__ import annotations

# --- FORMATO A (canonico): Cyberpunk 2077, app 1091500 --------------------
CYBERPUNK_MINIMUM = (
    '<strong>Minimum:</strong><br><ul class="bb_ul">'
    '<li>Requires a 64-bit processor and operating system<br></li>'
    '<li><strong>OS:</strong> 64-bit Windows 10<br></li>'
    '<li><strong>Processor:</strong> Core i7-6700 or Ryzen 5 1600<br></li>'
    '<li><strong>Memory:</strong> 12 GB RAM<br></li>'
    '<li><strong>Graphics:</strong> GeForce GTX 1060 6GB<br></li>'
    '<li><strong>DirectX:</strong> Version 12<br></li>'
    '<li><strong>Storage:</strong> 70 GB available space</li></ul>'
)

CYBERPUNK_RECOMMENDED = (
    '<strong>Recommended:</strong><br><ul class="bb_ul">'
    '<li>Requires a 64-bit processor and operating system<br></li>'
    '<li><strong>OS:</strong> 64-bit Windows 10<br></li>'
    '<li><strong>Processor:</strong> Core i7-12700 or Ryzen 7 7800X3D<br></li>'
    '<li><strong>Memory:</strong> 16 GB RAM<br></li>'
    '<li><strong>Graphics:</strong> GeForce RTX 2060 SUPER or Radeon RX 5700 XT '
    'or Arc A770<br></li>'
    '<li><strong>DirectX:</strong> Version 12<br></li>'
    '<li><strong>Storage:</strong> 70 GB available space</li></ul>'
)

# --- FORMATO B (rotulo E valor dentro do <strong>): Terraria, app 105600 ---
# Este e o caso que quebra parsers que assumem "valor = texto apos </strong>".
TERRARIA_RECOMMENDED = (
    '<h2 class="bb_tag" ><strong>RECOMMENDED</strong></h2><ul class="bb_ul">'
    '<li><strong>OS: Windows 7, 8/8.1, 10</strong> <br>\t\t\t\t</li>'
    '<li><strong>Processor: Dual Core 3.0 Ghz</strong> <br>\t\t\t\t</li>'
    '<li><strong>Memory: 4GB</strong><br>\t\t\t\t</li>'
    '<li><strong>Hard Disk Space: 200MB </strong>\t\t<br>\t\t\t\t</li>'
    '<li><strong>Video Card: 256mb Video Memory, capable of Shader Model 2.0+'
    '</strong></li></ul>'
)

# --- SO legado marcado com asterisco: Portal 2, app 620 -------------------
PORTAL2_MINIMUM = (
    '<strong>Minimum:</strong><br><ul class="bb_ul">'
    '<li><strong>OS *:</strong> Windows 7 / Vista / XP<br></li>'
    '<li><strong>Processor:</strong> 3.0 GHz P4, Dual Core 2.0 (or higher) '
    'or AMD64X2 (or higher)<br></li>'
    '<li><strong>Memory:</strong> 2 GB RAM<br></li>'
    '<li><strong>Graphics:</strong> Video card must be 128 MB or more<br></li>'
    '<li><strong>Hard Drive:</strong> At least 7.6 GB of free space</li></ul>'
)

# --- Hades: recomendado curto com OS legado, app 1145360 -----------------
HADES_RECOMMENDED = (
    '<strong>Recommended:</strong><br><ul class="bb_ul">'
    '<li><strong>OS *:</strong> Windows 7 SP1<br></li>'
    '<li><strong>Processor:</strong> Dual Core 3.0 GHz+<br></li>'
    '<li><strong>Memory:</strong> 8 GB RAM<br></li>'
    '<li><strong>Graphics:</strong> 2GB VRAM / DirectX 10+ support<br></li>'
    '<li><strong>Storage:</strong> 20 GB available space</li></ul>'
)


def appdetails(app_id: int, *, name: str, app_type: str = "game",
               release: str = "Dec 9, 2020", coming_soon: bool = False,
               windows: bool = True, mac: bool = False, linux: bool = False,
               minimum: str | None = None, recommended: str | None = None,
               reviews: int | None = 1000, genres: list[str] | None = None,
               success: bool = True, languages: str | None = None) -> dict:
    """Monta um payload de appdetails com a mesma forma da resposta real."""
    if not success:
        return {str(app_id): {"success": False}}
    pcr: dict[str, str] = {}
    if minimum:
        pcr["minimum"] = minimum
    if recommended:
        pcr["recommended"] = recommended
    return {
        str(app_id): {
            "success": True,
            "data": {
                "type": app_type,
                "name": name,
                "steam_appid": app_id,
                "release_date": {"coming_soon": coming_soon, "date": release},
                "platforms": {"windows": windows, "mac": mac, "linux": linux},
                "genres": [{"id": "3", "description": g}
                           for g in (genres or ["Action"])],
                "categories": [{"id": 2, "description": "Single-player"}],
                "developers": ["Dev Studio"],
                "publishers": ["Publisher Inc"],
                "recommendations": ({"total": reviews} if reviews is not None
                                    else {}),
                "supported_languages": languages or
                "English<strong>*</strong>, French, German",
                "is_free": False,
                "required_age": 0,
                "pc_requirements": pcr,
            },
        }
    }


def tags_payload(app_id: int, names: list[str] | None,
                 source: str = "STORE_HTML") -> dict:
    if names is None:
        return {"app_id": app_id, "tags": None, "tags_source": source,
                "collected_at": "2026-08-30T12:00:00Z"}
    return {
        "app_id": app_id,
        "tags": [{"tagid": 1000 + i, "name": n} for i, n in enumerate(names)],
        "tags_source": source,
        "collected_at": "2026-08-30T12:00:00Z",
    }
