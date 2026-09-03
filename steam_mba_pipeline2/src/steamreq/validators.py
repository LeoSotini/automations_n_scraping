"""Validacoes antes de considerar um registro concluido (item 8 do projeto).

Distincao central exigida pelo item 8, materializada em campos separados:

  collection_status  -> dado ausente na Steam | pagina inacessivel |
                        falha HTTP | falha de parsing
  exclusion_reason   -> item excluido pela METODOLOGIA

Nunca confundidos. Um jogo sem requisitos recomendados NAO e erro: gera
has_recommended_requirements=false e passa na validacao.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date
from typing import Any

from .config import Filters
from .logging_setup import get_logger

log = get_logger("validators")

SEVERITY_ERROR = "ERROR"
SEVERITY_WARNING = "WARNING"


@dataclass
class Issue:
    app_id: int | None
    code: str
    severity: str
    detail: str

    def as_dict(self) -> dict[str, Any]:
        return {"app_id": self.app_id, "code": self.code,
                "severity": self.severity, "detail": self.detail}


@dataclass
class ValidationReport:
    issues: list[Issue] = field(default_factory=list)
    checked: int = 0

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == SEVERITY_ERROR]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == SEVERITY_WARNING]

    @property
    def ok(self) -> bool:
        return not self.errors

    def add(self, app_id: int | None, code: str, severity: str,
            detail: str) -> None:
        self.issues.append(Issue(app_id, code, severity, detail))

    def counts_by_code(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for i in self.issues:
            out[i.code] = out.get(i.code, 0) + 1
        return dict(sorted(out.items(), key=lambda kv: -kv[1]))

    def as_dict(self) -> dict[str, Any]:
        return {"checked": self.checked, "ok": self.ok,
                "n_errors": len(self.errors), "n_warnings": len(self.warnings),
                "counts_by_code": self.counts_by_code(),
                "issues": [i.as_dict() for i in self.issues[:1000]]}


REQUIRED_KEYS = (
    "app_id", "name", "steam_url", "type", "release_date_raw", "release_date",
    "release_year", "genres", "categories", "steam_tags", "developer",
    "publisher", "review_count", "platform_windows", "pc_requirements",
    "has_minimum_requirements", "has_recommended_requirements",
    "included_initially", "exclusion_reason", "filter_version",
    "scraper_version", "collected_at", "source", "collection_status",
)


def validate_record(record: dict[str, Any], filters: Filters,
                    report: ValidationReport) -> None:
    app_id = record.get("app_id")

    # app_id valido
    if not isinstance(app_id, int) or app_id <= 0:
        report.add(app_id, "INVALID_APP_ID", SEVERITY_ERROR,
                   f"app_id invalido: {app_id!r}")
        return

    # homogeneidade do schema (item 6)
    missing = [k for k in REQUIRED_KEYS if k not in record]
    if missing:
        report.add(app_id, "MISSING_SCHEMA_KEYS", SEVERITY_ERROR,
                   f"chaves ausentes no registro: {missing}")

    included = bool(record.get("included_initially"))

    # coerencia entre inclusao e motivo (item 2.17)
    if included and record.get("exclusion_reason") is not None:
        report.add(app_id, "INCLUDED_WITH_REASON", SEVERITY_ERROR,
                   "registro incluido mas com exclusion_reason preenchido")
    if not included and not record.get("exclusion_reason"):
        report.add(app_id, "EXCLUDED_WITHOUT_REASON", SEVERITY_ERROR,
                   "registro excluido sem exclusion_reason (viola o item 2.17)")

    # separacao rigorosa entre minimos e recomendados (item 5)
    pcr = record.get("pc_requirements") or {}
    minimum = pcr.get("minimum") or {}
    recommended = pcr.get("recommended") or {}
    if not isinstance(minimum, dict) or not isinstance(recommended, dict):
        report.add(app_id, "REQUIREMENTS_NOT_SEPARATED", SEVERITY_ERROR,
                   "pc_requirements sem os blocos minimum/recommended")
    else:
        has_rec = bool(record.get("has_recommended_requirements"))
        if has_rec and recommended.get("raw") is None:
            report.add(app_id, "RECOMMENDED_FLAG_MISMATCH", SEVERITY_ERROR,
                       "has_recommended_requirements=true mas raw ausente")
        if not has_rec and recommended.get("raw") is not None:
            report.add(app_id, "RECOMMENDED_FLAG_MISMATCH", SEVERITY_ERROR,
                       "has_recommended_requirements=false mas raw presente")
        # O texto bruto e o dado insubstituivel (item 5).
        if record.get("has_minimum_requirements") and minimum.get("raw") is None:
            report.add(app_id, "MINIMUM_RAW_MISSING", SEVERITY_ERROR,
                       "has_minimum_requirements=true mas raw do minimo ausente")
        # Contaminacao: recomendado identico ao minimo E com raw identico
        # indica copia indevida (o parser jamais copia; isto e uma rede de
        # seguranca contra regressao).
        if (minimum.get("raw") is not None
                and minimum.get("raw") == recommended.get("raw")
                and not has_rec):
            report.add(app_id, "MIN_LEAKED_INTO_RECOMMENDED", SEVERITY_ERROR,
                       "raw do minimo aparece como recomendado sem a flag")
        if recommended.get("unparsed_labels"):
            report.add(app_id, "UNPARSED_LABELS", SEVERITY_WARNING,
                       f"rotulos nao mapeados no recomendado: "
                       f"{[u['label'] for u in recommended['unparsed_labels']][:5]}")

    if included:
        # tipo compativel com jogo (item 2.1)
        if record.get("type") not in filters.app_type_allowed:
            report.add(app_id, "INCLUDED_WRONG_TYPE", SEVERITY_ERROR,
                       f"incluido com type={record.get('type')!r}")
        # plataforma (item 2.2)
        if filters.require_windows and not record.get("platform_windows"):
            report.add(app_id, "INCLUDED_WITHOUT_WINDOWS", SEVERITY_ERROR,
                       "incluido sem suporte a Windows")
        # data consistente e ano >= limite (item 2.3)
        rd = record.get("release_date")
        if rd is None:
            report.add(app_id, "INCLUDED_WITHOUT_DATE", SEVERITY_WARNING,
                       f"incluido sem data utilizavel "
                       f"(raw={record.get('release_date_raw')!r})")
        else:
            try:
                parsed = date.fromisoformat(rd)
            except (TypeError, ValueError):
                report.add(app_id, "INVALID_DATE_FORMAT", SEVERITY_ERROR,
                           f"release_date nao e ISO: {rd!r}")
            else:
                if rd < filters.min_date:
                    report.add(app_id, "INCLUDED_BEFORE_MIN_DATE", SEVERITY_ERROR,
                               f"incluido com release_date={rd} < "
                               f"{filters.min_date}")
                if record.get("release_year") != parsed.year:
                    report.add(app_id, "YEAR_DATE_MISMATCH", SEVERITY_ERROR,
                               f"release_year={record.get('release_year')} "
                               f"incoerente com {rd}")
                if parsed.year > date.today().year + 1:
                    report.add(app_id, "IMPLAUSIBLE_FUTURE_DATE", SEVERITY_WARNING,
                               f"data futura implausivel: {rd}")
        # filtro de tags executado (item 8)
        if record.get("steam_tags") is None and \
                record.get("tags_source") not in ("AGE_GATE_LOGIN_REQUIRED",
                                                  "AGE_GATE_INTERSTITIAL",
                                                  "UNAVAILABLE", "SKIPPED"):
            report.add(app_id, "TAGS_NOT_EVALUATED", SEVERITY_ERROR,
                       "incluido sem tags e sem motivo declarado em tags_source")
        # o item 2.10 exclui Indie na fonte; um incluido com is_indie sinaliza
        # que o untags do discovery falhou
        if record.get("is_indie"):
            report.add(app_id, "INDIE_PASSED_SOURCE_FILTER", SEVERITY_WARNING,
                       "incluido com is_indie=true: o untags do discovery pode "
                       "ter falhado, ou o app veio de outra origem")
        # item 2.8: conflito deve estar sinalizado para revisao
        if record.get("tag_conflict_3d_2d") and not record.get("needs_manual_review"):
            report.add(app_id, "CONFLICT_NOT_FLAGGED", SEVERITY_ERROR,
                       "conflito 3D+2D sem needs_manual_review=true")

    # serializavel em JSON / UTF-8 (item 8)
    try:
        json.dumps(record, ensure_ascii=False)
    except (TypeError, ValueError) as exc:
        report.add(app_id, "NOT_JSON_SERIALIZABLE", SEVERITY_ERROR, str(exc))


def validate_dataset(records: list[dict[str, Any]],
                     filters: Filters) -> ValidationReport:
    report = ValidationReport()
    seen: dict[int, int] = {}
    for record in records:
        report.checked += 1
        validate_record(record, filters, report)
        aid = record.get("app_id")
        if isinstance(aid, int):
            seen[aid] = seen.get(aid, 0) + 1
    for aid, n in seen.items():
        if n > 1:
            report.add(aid, "DUPLICATE_APP_ID", SEVERITY_ERROR,
                       f"app_id repetido {n} vezes")
    return report
