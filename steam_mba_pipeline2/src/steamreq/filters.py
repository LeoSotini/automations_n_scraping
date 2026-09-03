"""Filtros metodologicos (item 2 do projeto) e montagem do registro final.

Principios que governam este modulo:

* Item 2.16 — PRIORIZAR RECALL. Em duvida, INCLUI e marca needs_manual_review.
* Item 2.17 — nenhum descarte sem exclusion_reason registrado.
* Item 2.8 — conflito 3D+2D nao exclui (1.915 casos medidos na FASE 1).
* Item 2.9 — "2.5D" nunca exclui.
* Item 2.11 — negativas candidatas apenas MARCAM (2.389 jogos sao 3D e
  Pixel Graphics simultaneamente).
* Item 2.12 — Casual / Free to Play / Early Access nunca excluem.
* Item 5 — minimum e recommended jamais se misturam.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from .config import Filters
from .logging_setup import get_logger, utc_now_iso
from .requirements_parser import parse_pc_requirements

log = get_logger("filters")

STORE_URL = "https://store.steampowered.com/app/{app_id}/"


class ExclusionReason:
    """Motivos de exclusao (item 2.17). Ordem de avaliacao = item 2.18."""

    NOT_GAME = "NOT_GAME"
    NO_WINDOWS_SUPPORT = "NO_WINDOWS_SUPPORT"
    UNRELEASED = "UNRELEASED"
    BEFORE_MIN_DATE = "BEFORE_2005"
    NO_3D_EVIDENCE = "NO_3D_EVIDENCE"
    TWO_DIMENSIONAL = "TWO_DIMENSIONAL"
    BELOW_REVIEW_THRESHOLD = "BELOW_REVIEW_THRESHOLD"
    INDIE = "INDIE"
    NO_DATA = "NO_DATA"
    INVALID_RELEASE_DATE = "INVALID_RELEASE_DATE"


class InclusionBasis:
    STRONG_3D_TAG = "STRONG_3D_TAG"
    SECONDARY_ONLY = "SECONDARY_ONLY"
    TAGS_UNAVAILABLE = "TAGS_UNAVAILABLE"


_MONTHS = {m.casefold(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"], start=1)}
_MONTHS.update({m[:3].casefold(): i for m, i in list(_MONTHS.items())})


def parse_release_date(raw: str | None) -> tuple[date | None, bool]:
    """Converte a data de lancamento da Steam para `date`.

    Retorna (data, é_precisa). Formatos observados/possiveis:
      "Dec 9, 2020" | "9 Dec, 2020" | "December 2020" | "2020" |
      "Q1 2021" | "Coming soon" | ""

    Quando so o ano/mes e conhecido, assume o primeiro dia do periodo e marca
    é_precisa=False. O item 2.3 exige preservar a data ORIGINAL — por isso
    release_date_raw e sempre gravado sem alteracao.
    """
    if not raw or not str(raw).strip():
        return None, False
    text = str(raw).strip()

    # "Dec 9, 2020" / "December 9, 2020"
    m = re.match(r"^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s*(\d{4})$", text)
    if m and m.group(1).casefold() in _MONTHS:
        try:
            return date(int(m.group(3)), _MONTHS[m.group(1).casefold()],
                        int(m.group(2))), True
        except ValueError:
            return None, False

    # "9 Dec, 2020" / "9 December 2020"
    m = re.match(r"^(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s*(\d{4})$", text)
    if m and m.group(2).casefold() in _MONTHS:
        try:
            return date(int(m.group(3)), _MONTHS[m.group(2).casefold()],
                        int(m.group(1))), True
        except ValueError:
            return None, False

    # "December 2020" / "Dec 2020"
    m = re.match(r"^([A-Za-z]{3,9})\.?\s+(\d{4})$", text)
    if m and m.group(1).casefold() in _MONTHS:
        return date(int(m.group(2)), _MONTHS[m.group(1).casefold()], 1), False

    # "Q1 2021"
    m = re.match(r"^Q([1-4])\s+(\d{4})$", text, re.IGNORECASE)
    if m:
        return date(int(m.group(2)), (int(m.group(1)) - 1) * 3 + 1, 1), False

    # "2020"
    m = re.match(r"^(\d{4})$", text)
    if m:
        return date(int(m.group(1)), 1, 1), False

    # ISO, por robustez
    try:
        return datetime.strptime(text, "%Y-%m-%d").date(), True
    except ValueError:
        pass

    return None, False


def count_supported_languages(raw: str | None) -> int | None:
    """Conta idiomas na string HTML de supported_languages.

    Proxy auxiliar de escala de producao (item 2.15). Sem interpretacao
    adicional: apenas conta as entradas separadas por virgula.
    """
    if not raw:
        return None
    text = re.sub(r"<[^>]+>", "", str(raw))
    text = re.sub(r"\blanguages with full audio support\b", "", text,
                  flags=re.IGNORECASE)
    parts = [p.strip(" *\u00a0") for p in text.split(",")]
    return len([p for p in parts if p]) or None


@dataclass
class TagEvaluation:
    """Resultado da avaliacao de tags (itens 2.5 a 2.11)."""

    matched_positive_strong: list[str]
    matched_positive_secondary: list[str]
    matched_2d: list[str]
    matched_neutral: list[str]
    negative_candidate_tags: list[str]
    is_indie: bool
    tags_available: bool

    @property
    def has_2d_tags(self) -> bool:
        return bool(self.matched_2d)

    @property
    def has_strong(self) -> bool:
        return bool(self.matched_positive_strong)

    @property
    def has_secondary(self) -> bool:
        return bool(self.matched_positive_secondary)

    @property
    def tag_conflict_3d_2d(self) -> bool:
        """Item 2.8: sinais contraditorios. Preserva em vez de descartar."""
        return self.has_strong and self.has_2d_tags

    @property
    def matched_positive_tags(self) -> list[str]:
        """Item 2.6: preservar quais tags produziram a inclusao."""
        return self.matched_positive_strong + self.matched_positive_secondary


def evaluate_tags(tag_names: list[str] | None, filters: Filters) -> TagEvaluation:
    if tag_names is None:
        return TagEvaluation([], [], [], [], [], False, tags_available=False)

    present = {t.strip().casefold() for t in tag_names}

    def hits(bucket: list[str]) -> list[str]:
        return [t for t in bucket if t.strip().casefold() in present]

    return TagEvaluation(
        matched_positive_strong=hits(filters.positive_strong),
        matched_positive_secondary=hits(filters.positive_secondary),
        matched_2d=hits(filters.negative_2d),
        matched_neutral=hits(filters.neutral_never_exclude),
        negative_candidate_tags=hits(filters.negative_candidates),
        is_indie=bool(hits(filters.exclude_at_source)),
        tags_available=True,
    )


def _decide(record: dict[str, Any], ev: TagEvaluation,
            filters: Filters) -> tuple[bool, str | None, str | None, bool]:
    """Aplica os criterios na ordem do item 2.18.

    Retorna (incluido, exclusion_reason, inclusion_basis, needs_manual_review).
    O PRIMEIRO criterio reprovado define o motivo — garante motivo unico e
    deterministico por registro.
    """
    # 2.1 tipo de app
    if record.get("type") not in filters.app_type_allowed:
        return False, ExclusionReason.NOT_GAME, None, False

    # 2.2 plataforma
    if filters.require_windows and not record.get("platform_windows"):
        return False, ExclusionReason.NO_WINDOWS_SUPPORT, None, False

    # 2.4 estado de lancamento
    if filters.exclude_coming_soon and record.get("coming_soon"):
        return False, ExclusionReason.UNRELEASED, None, False

    # 2.3 data de lancamento
    rd = record.get("release_date")
    if rd is None:
        # Sem data utilizavel: recall sobre precisao (2.16) — preserva para
        # revisao manual em vez de descartar silenciosamente.
        return True, None, InclusionBasis.TAGS_UNAVAILABLE if not ev.tags_available \
            else (InclusionBasis.STRONG_3D_TAG if ev.has_strong
                  else InclusionBasis.SECONDARY_ONLY), True
    if rd < filters.min_date:
        return False, ExclusionReason.BEFORE_MIN_DATE, None, False

    # 2.5 a 2.11 tags
    if not ev.tags_available:
        # Tags inacessiveis (ex.: age gate por login). Nao ha base para excluir
        # por tag; preserva com revisao manual (2.16).
        return True, None, InclusionBasis.TAGS_UNAVAILABLE, True

    if ev.has_strong:
        # Item 2.8: conflito 3D+2D com evidencia FORTE de 3D. Sinais realmente
        # contraditorios -> INCLUI e marca para resolucao posterior.
        return True, None, InclusionBasis.STRONG_3D_TAG, ev.tag_conflict_3d_2d

    if ev.has_secondary:
        # Decisao D10: sem nenhuma tag forte de 3D, a evidencia negativa do
        # item 2.8 prevalece sobre a evidencia complementar do item 2.7.
        # Motivo: as secundarias nao atestam tridimensionalidade (o proprio
        # item 2.7 diz isso), e medimos que "Action-Adventure" tem 3.820 jogos
        # 2D nao-Indie. Terraria e Hollow Knight entravam por aqui.
        if ev.has_2d_tags:
            return False, ExclusionReason.TWO_DIMENSIONAL, None, False
        # Item 2.7 + 2.16: nao garante 3D, mas preserva com revisao manual.
        return True, None, InclusionBasis.SECONDARY_ONLY, True

    # Sem nenhuma evidencia positiva. Se houver 2D, o motivo mais informativo
    # e a bidimensionalidade (2.8); caso contrario, ausencia de evidencia 3D.
    if ev.has_2d_tags:
        return False, ExclusionReason.TWO_DIMENSIONAL, None, False
    return False, ExclusionReason.NO_3D_EVIDENCE, None, False


def build_record(app_id: int, appdetails_payload: dict | None,
                 tags_payload: dict | None, filters: Filters, *,
                 scraper_version: str,
                 collected_at: str | None = None) -> dict[str, Any]:
    """Monta o registro homogeneo do schema a partir dos payloads brutos.

    Opera EXCLUSIVAMENTE sobre data/raw — nunca sobre a rede. Reexecutar este
    estagio com filtros diferentes nao custa nenhuma requisicao (itens 5 e 12).
    """
    node = (appdetails_payload or {}).get(str(app_id)) or {}
    success = bool(node.get("success"))
    data = node.get("data") or {}

    tags_source = (tags_payload or {}).get("tags_source") or "UNAVAILABLE"
    tag_items = (tags_payload or {}).get("tags")
    tag_names = [t["name"] for t in tag_items] if tag_items else None
    tag_ids = [t["tagid"] for t in tag_items] if tag_items else None

    release_raw = (data.get("release_date") or {}).get("date")
    release_dt, date_precise = parse_release_date(release_raw)
    platforms = data.get("platforms") or {}

    record: dict[str, Any] = {
        # identificacao
        "app_id": app_id,
        "name": data.get("name"),
        "steam_url": STORE_URL.format(app_id=app_id),
        "type": data.get("type"),
        # lancamento (2.3: data original SEMPRE preservada)
        "release_date_raw": release_raw,
        "release_date": release_dt.isoformat() if release_dt else None,
        "release_year": release_dt.year if release_dt else None,
        "release_date_precise": date_precise,
        "coming_soon": bool((data.get("release_date") or {}).get("coming_soon")),
        # classificacao
        "genres": [g.get("description") for g in (data.get("genres") or [])],
        "categories": [c.get("description") for c in (data.get("categories") or [])],
        "steam_tags": tag_names,
        "steam_tag_ids": tag_ids,
        "tags_source": tags_source,
        "developer": data.get("developers"),
        "publisher": data.get("publishers"),
        # metadados auxiliares
        "review_count": (data.get("recommendations") or {}).get("total"),
        "review_count_metric": filters.review_metric,
        "metacritic_score": (data.get("metacritic") or {}).get("score"),
        "supported_languages_raw": data.get("supported_languages"),
        "supported_languages_count": count_supported_languages(
            data.get("supported_languages")),
        "is_free": data.get("is_free"),
        "required_age": data.get("required_age"),
        "platform_windows": bool(platforms.get("windows")),
        "platform_mac": bool(platforms.get("mac")),
        "platform_linux": bool(platforms.get("linux")),
    }

    # requisitos (item 5) — minimum e recommended independentes
    record.update(parse_pc_requirements(data.get("pc_requirements")))

    ev = evaluate_tags(tag_names, filters)

    if not success or not data:
        included, reason, basis, review = False, ExclusionReason.NO_DATA, None, False
    else:
        included, reason, basis, review = _decide(record, ev, filters)

        # 2.14: threshold de reviews, apenas se configurado.
        if included and filters.review_min_threshold is not None:
            rc = record.get("review_count")
            if rc is None or int(rc) < int(filters.review_min_threshold):
                included, reason, basis = (
                    False, ExclusionReason.BELOW_REVIEW_THRESHOLD, None)

    record.update({
        "included_initially": included,
        "exclusion_reason": reason,
        "inclusion_basis": basis,
        "matched_positive_tags": ev.matched_positive_tags,
        "matched_positive_strong": ev.matched_positive_strong,
        "matched_positive_secondary": ev.matched_positive_secondary,
        "is_indie": ev.is_indie,
        "has_2d_tags": ev.has_2d_tags,
        "matched_2d_tags": ev.matched_2d,
        "tag_conflict_3d_2d": ev.tag_conflict_3d_2d,
        "neutral_tags": ev.matched_neutral,
        "negative_candidate_tags": ev.negative_candidate_tags,
        "needs_manual_review": review,
        # reprodutibilidade (item 12)
        "filter_version": filters.filter_version,
        "scraper_version": scraper_version,
        "collected_at": collected_at or (tags_payload or {}).get("collected_at")
        or utc_now_iso(),
        "source": "steam_store_api",
        "source_urls": {
            "appdetails": "https://store.steampowered.com/api/appdetails"
                          f"?appids={app_id}&l=english&cc=us",
            "store_page": STORE_URL.format(app_id=app_id),
        },
        "collection_status": (
            "FAILED" if not success
            else "COMPLETE" if tag_names else "PARTIAL_NO_TAGS"),
    })
    return record


def summarize_exclusions(records: list[dict[str, Any]]) -> dict[str, int]:
    out: dict[str, int] = {}
    for r in records:
        if not r.get("included_initially"):
            key = r.get("exclusion_reason") or "UNKNOWN"
            out[key] = out.get(key, 0) + 1
    return out
