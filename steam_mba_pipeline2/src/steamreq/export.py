"""Estagios 3, 4 e 6 — filtro offline, parsing e export (itens 6 e 12).

Opera EXCLUSIVAMENTE sobre data/raw. Nenhuma requisicao de rede. Consequencia:
alterar filtros, corrigir o parser ou mover o corte de 2005 nao custa uma
unica requisicao a Steam.
"""
from __future__ import annotations

import os
from typing import Any

from .config import Filters, Settings
from .filters import build_record, summarize_exclusions
from .logging_setup import get_logger, utc_now_iso
from .requirements_parser import FIELDS
from .storage import RawStore, append_ledger, read_json, write_json_atomic

log = get_logger("export")


def build_all_records(settings: Settings, filters: Filters, *,
                      app_ids: list[int] | None = None) -> list[dict[str, Any]]:
    """Reconstroi todos os registros a partir dos payloads brutos."""
    raw = RawStore(settings.path("raw_appdetails"), settings.path("raw_tags"))
    ids = app_ids if app_ids is not None else raw.known_app_ids()
    records: list[dict[str, Any]] = []
    for app_id in ids:
        payload = raw.load_appdetails(app_id)
        if payload is None:
            continue
        tags_payload = raw.load_tags(app_id)
        records.append(build_record(app_id, payload, tags_payload, filters,
                                   scraper_version=settings.scraper_version))
    return records


def run_filter_stage(settings: Settings, filters: Filters) -> dict[str, Any]:
    """Aplica a metodologia e grava o ledger de exclusoes (item 2.17)."""
    records = build_all_records(settings, filters)
    total = len(records)
    log.info("[FILTER]    entrada %s registros (a partir de data/raw)",
             f"{total:,}")

    summary = summarize_exclusions(records)
    remaining = total
    # Reporta na ordem de avaliacao do item 2.18, para leitura de funil.
    order = ["NOT_GAME", "NO_WINDOWS_SUPPORT", "UNRELEASED", "BEFORE_2005",
             "NO_3D_EVIDENCE", "TWO_DIMENSIONAL", "BELOW_REVIEW_THRESHOLD",
             "NO_DATA"]
    for reason in order:
        n = summary.get(reason, 0)
        if n:
            remaining -= n
            log.info("[FILTER]      %-22s -> -%-7s (%s restantes)",
                     reason, f"{n:,}", f"{remaining:,}")
    for reason, n in summary.items():
        if reason not in order:
            remaining -= n
            log.info("[FILTER]      %-22s -> -%-7s (%s restantes)",
                     reason, f"{n:,}", f"{remaining:,}")

    included = [r for r in records if r.get("included_initially")]
    conflicts = [r for r in included if r.get("tag_conflict_3d_2d")]
    review = [r for r in included if r.get("needs_manual_review")]
    no_rec = [r for r in included if not r.get("has_recommended_requirements")]

    log.info("[FILTER]    conflitos 3D+2D preservados: %s "
             "(needs_manual_review=true)", f"{len(conflicts):,}")
    log.info("[FILTER]    marcados para revisao manual: %s", f"{len(review):,}")
    log.info("[FILTER]    incluidos SEM requisitos recomendados: %s (%.1f%%)",
             f"{len(no_rec):,}",
             100.0 * len(no_rec) / len(included) if included else 0.0)
    log.info("[FILTER]    saida %s registros incluidos", f"{len(included):,}")

    excluded_records = [
        {"app_id": r["app_id"], "name": r.get("name"),
         "release_date": r.get("release_date"),
         "release_date_raw": r.get("release_date_raw"),
         "type": r.get("type"), "platform_windows": r.get("platform_windows"),
         "steam_tags": r.get("steam_tags"), "review_count": r.get("review_count"),
         "included_initially": False,
         "exclusion_reason": r.get("exclusion_reason"),
         "filter_version": r.get("filter_version")}
        for r in records if not r.get("included_initially")]
    ledger_path = f"{settings.path('processed')}/ledger_exclusions.json"
    append_ledger(ledger_path, excluded_records,
                  filter_version=filters.filter_version)
    log.info("[FILTER]    ledger de exclusoes -> %s (%s registros)",
             ledger_path, f"{len(excluded_records):,}")

    write_json_atomic(f"{settings.path('processed')}/filtered_records.json",
                      records, indent=1)
    return {"records": records, "included": included, "summary": summary}


def flatten_record(record: dict[str, Any]) -> dict[str, Any]:
    """Achata pc_requirements em minimum_* / recommended_* (item 5).

    Nomes inequivocamente diferentes, como o projeto exige:
      minimum_cpu, minimum_ram, ...  |  recommended_cpu, recommended_ram, ...
    """
    flat = {k: v for k, v in record.items() if k != "pc_requirements"}
    pcr = record.get("pc_requirements") or {}
    for block_name in ("minimum", "recommended"):
        block = pcr.get(block_name) or {}
        flat[f"{block_name}_raw"] = block.get("raw")
        for f in FIELDS:
            flat[f"{block_name}_{f}"] = block.get(f)
        flat[f"{block_name}_os_legacy_flag"] = block.get("os_legacy_flag", False)
        flat[f"{block_name}_unparsed_labels"] = block.get("unparsed_labels") or []
    return flat


def export_dataset(settings: Settings, filters: Filters, *, flat: bool = False,
                   include_excluded: bool = False) -> str:
    """Grava o dataset final como JSON tabular (item 6)."""
    stage = run_filter_stage(settings, filters)
    records = stage["records"] if include_excluded else stage["included"]

    if flat:
        records = [flatten_record(r) for r in records]

    # Homogeneidade: todo registro recebe todas as chaves observadas (item 6).
    all_keys: dict[str, None] = {}
    for r in records:
        for k in r:
            all_keys.setdefault(k, None)
    normalized = [{k: r.get(k) for k in all_keys} for r in records]

    out_path = os.path.join(settings.path("processed"), "dataset.json")
    write_json_atomic(out_path, normalized, indent=1)

    meta = {
        "generated_at": utc_now_iso(),
        "scraper_version": settings.scraper_version,
        "filter_version": filters.filter_version,
        "source": "steam_store_api",
        "format": "flat" if flat else "nested",
        "n_records": len(normalized),
        "includes_excluded": include_excluded,
        "review_count_metric": filters.review_metric,
        "review_min_threshold": filters.review_min_threshold,
        "exclusion_summary": stage["summary"],
        "columns": list(all_keys),
        "known_limitations": [
            "discovery via storefront: jogos deslistados ausentes (vies de "
            "sobrevivencia correlacionado ao tempo)",
            "tags Steam sao atribuidas por usuarios e refletem o estado ATUAL, "
            "nao o do lancamento",
            "requisitos refletem a versao ATUAL da pagina, nao a do lancamento",
            "review_count usa appdetails.recommendations.total, que difere de "
            "appreviews.total_reviews",
        ],
    }
    meta_path = os.path.join(settings.path("processed"), "dataset_metadata.json")
    write_json_atomic(meta_path, meta, indent=2)

    size = os.path.getsize(out_path)
    log.info("[EXPORT]    %s registros -> %s (UTF-8, %.1f MB, formato %s)",
             f"{len(normalized):,}", out_path, size / 1024 / 1024,
             "flat" if flat else "nested")
    log.info("[EXPORT]    metadados -> %s", meta_path)
    return out_path


def analyze_thresholds(settings: Settings, filters: Filters) -> dict[str, Any]:
    """Item 2.14: produz a evidencia antes de fechar o threshold de reviews.

    Deliberadamente NAO aplica nenhum corte. Apenas mede, para que a decisao
    seja justificavel e reproduzivel.
    """
    records = build_all_records(settings, filters)
    # Avalia sobre quem passou nos demais criterios, isolando o efeito do corte.
    base = [r for r in records if r.get("included_initially")
            or r.get("exclusion_reason") == "BELOW_REVIEW_THRESHOLD"]
    counts = sorted(int(r.get("review_count") or 0) for r in base)
    n = len(counts)

    def pct(p: float) -> int | None:
        if not counts:
            return None
        return counts[min(n - 1, max(0, int(round(p / 100 * (n - 1)))))]

    survivors = {}
    for t in filters.review_candidate_thresholds:
        kept = [c for c in counts if c >= t]
        survivors[str(t)] = {
            "n": len(kept),
            "pct_of_base": round(100.0 * len(kept) / n, 2) if n else 0.0,
        }

    result = {
        "generated_at": utc_now_iso(),
        "metric": filters.review_metric,
        "note": "nenhum corte aplicado; este relatorio existe para cumprir os "
                "4 passos exigidos pelo item 2.14 antes de fechar o valor",
        "base_population": n,
        "distribution": {
            "min": counts[0] if counts else None,
            "p10": pct(10), "p25": pct(25), "median": pct(50),
            "p75": pct(75), "p90": pct(90), "p99": pct(99),
            "max": counts[-1] if counts else None,
            "zero_reviews": sum(1 for c in counts if c == 0),
        },
        "survivors_by_threshold": survivors,
    }
    out = os.path.join(settings.path("processed"), "threshold_analysis.json")
    write_json_atomic(out, result, indent=2)

    log.info("[THRESHOLD] populacao base: %s registros", f"{n:,}")
    d = result["distribution"]
    log.info("[THRESHOLD] distribuicao de %s: min=%s p25=%s mediana=%s p75=%s "
             "p90=%s max=%s | zero reviews=%s", filters.review_metric,
             d["min"], d["p25"], d["median"], d["p75"], d["p90"], d["max"],
             f"{d['zero_reviews']:,}")
    for t, info in survivors.items():
        log.info("[THRESHOLD]   >= %-6s -> %s sobreviventes (%.1f%% da base)",
                 t, f"{info['n']:,}", info["pct_of_base"])
    log.info("[THRESHOLD] relatorio -> %s", out)
    log.info("[THRESHOLD] o item 2.14 exige ainda inspecao manual de falsos "
             "positivos/negativos antes de fixar o corte em filters.yaml")
    return result


def load_dataset(settings: Settings) -> list[dict[str, Any]]:
    data = read_json(os.path.join(settings.path("processed"), "dataset.json"),
                     default=[])
    return data if isinstance(data, list) else data.get("records", [])
