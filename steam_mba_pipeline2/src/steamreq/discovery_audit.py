"""Auditoria do viés de sobrevivência (decisao Q1 / D1).

Problema metodologico identificado na FASE 1: o /search/results/ retorna apenas
jogos ATUALMENTE visiveis na storefront. Jogos deslistados desaparecem, e a
probabilidade de deslistagem cresce com a idade do titulo. Isso produz vies de
sobrevivencia CORRELACIONADO COM O TEMPO — exatamente o eixo do estudo
longitudinal. Efeito esperado: enviesar PARA CIMA a taxa de crescimento
estimada dos requisitos, porque os jogos antigos que sobrevivem na loja tendem
a ser os de maior orcamento.

Este modulo quantifica a magnitude do vies comparando o catalogo enumerado por
app_id (IStoreService/GetAppList, exige API key) com o resultado do discovery.

NAO e um estagio obrigatorio: sem STEAM_API_KEY, e ignorado com WARNING e o
pipeline principal roda normalmente.
"""
from __future__ import annotations

import os
from typing import Any

from .config import Settings
from .logging_setup import fmt_elapsed, get_logger, utc_now_iso
from .storage import read_json, write_json_atomic

log = get_logger("audit")

API_KEY_ENV = "STEAM_API_KEY"


class MissingAPIKey(RuntimeError):
    pass


def get_api_key() -> str:
    key = os.environ.get(API_KEY_ENV, "").strip()
    if not key:
        raise MissingAPIKey(
            f"variavel de ambiente {API_KEY_ENV} nao definida. Gere uma chave "
            "gratuita em https://steamcommunity.com/dev/apikey e defina "
            f"{API_KEY_ENV}. A auditoria de vies e opcional; o pipeline "
            "principal funciona sem ela."
        )
    return key


def fetch_catalog(client, settings: Settings, *,  # noqa: ANN001
                  max_pages: int = 200) -> dict[int, str]:
    """Enumera o catalogo de jogos por app_id, paginando por last_appid."""
    key = get_api_key()
    catalog: dict[int, str] = {}
    last_appid: int | None = None
    for page in range(1, max_pages + 1):
        params: dict[str, Any] = {
            "key": key, "include_games": "true", "include_dlc": "false",
            "include_software": "false", "include_videos": "false",
            "include_hardware": "false", "max_results": 50000,
        }
        if last_appid is not None:
            params["last_appid"] = last_appid
        payload = client.get_json("api", settings.endpoints["api_app_list"], params)
        response = payload.get("response") or {}
        apps = response.get("apps") or []
        for app in apps:
            aid = app.get("appid")
            if aid is not None:
                catalog[int(aid)] = str(app.get("name") or "")
        log.info("[AUDIT]     pagina %d | +%s apps | catalogo=%s | decorrido %s",
                 page, f"{len(apps):,}", f"{len(catalog):,}", fmt_elapsed())
        if not response.get("have_more_results"):
            break
        last_appid = response.get("last_appid") or (
            apps[-1].get("appid") if apps else None)
        if last_appid is None:
            break
    return catalog


def run_audit(client, settings: Settings) -> dict[str, Any]:  # noqa: ANN001
    """Compara catalogo vs. discovery e estratifica a diferenca por ano.

    A estratificacao por ano usa os registros ja coletados em dataset.json,
    quando existirem — e a unica fonte de release_date disponivel offline.
    """
    processed = settings.path("processed")
    candidates = read_json(f"{processed}/candidates.json", default=None) or {}
    cand_ids = {int(c["app_id"]) for c in candidates.get("candidates", [])}
    if not cand_ids:
        log.warning("[AUDIT]     candidates.json vazio; rode `discover` antes "
                    "para que a comparacao seja significativa")

    catalog = fetch_catalog(client, settings)
    log.info("[AUDIT]     catalogo via API = %s apps | discovery = %s candidatos",
             f"{len(catalog):,}", f"{len(cand_ids):,}")

    dataset = read_json(f"{processed}/dataset.json", default=None) or []
    if isinstance(dataset, dict):
        dataset = dataset.get("records", [])
    year_by_id = {int(r["app_id"]): r.get("release_year")
                  for r in dataset if r.get("app_id") is not None}

    not_in_discovery = sorted(set(catalog) - cand_ids)
    not_in_catalog = sorted(cand_ids - set(catalog))

    by_year: dict[str, dict[str, int]] = {}
    for aid in cand_ids:
        year = year_by_id.get(aid)
        if year:
            slot = by_year.setdefault(str(year), {"in_discovery": 0,
                                                  "in_catalog": 0})
            slot["in_discovery"] += 1
            if aid in catalog:
                slot["in_catalog"] += 1

    result = {
        "meta": {
            "generated_at": utc_now_iso(),
            "method": "IStoreService/GetAppList vs store_search_results",
            "purpose": "quantificar vies de sobrevivencia do discovery via "
                       "storefront (decisao D1)",
        },
        "catalog_size": len(catalog),
        "discovery_size": len(cand_ids),
        "in_catalog_not_in_discovery": len(not_in_discovery),
        "in_discovery_not_in_catalog": len(not_in_catalog),
        "by_release_year": dict(sorted(by_year.items())),
        "sample_in_catalog_not_in_discovery": [
            {"app_id": a, "name": catalog[a]} for a in not_in_discovery[:200]],
    }
    out_path = f"{processed}/bias_audit.json"
    write_json_atomic(out_path, result, indent=2)
    log.info("[AUDIT]     no catalogo mas ausentes do discovery: %s",
             f"{len(not_in_discovery):,}")
    log.info("[AUDIT]     no discovery mas ausentes do catalogo: %s",
             f"{len(not_in_catalog):,}")
    log.info("[AUDIT]     resultado -> %s", out_path)
    return result
