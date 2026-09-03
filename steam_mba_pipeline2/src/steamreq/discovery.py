"""Estagio 1 — DISCOVERY via /search/results/ (item 2.18).

Fatos empiricos da FASE 1 que ditam a implementacao:

* `tags` com multiplos valores e AND, NAO OR (medido: 3D=51.760, FPS=10.988,
  "3D,FPS"=6.503). Logo a UNIAO das tags positivas do item 2.6 precisa ser
  feita no cliente: uma query por tag.
* `untags` e complemento exato (validado: 25.036 + 26.724 = 51.760). E usado
  para excluir "Indie" na fonte (decisao Q2/D2).
* `category1=998` filtra jogos no servidor; `os=win` exige Windows.
* Paginacao profunda funciona (start testado ate 45.000); count max = 50.
* Rate limit severo: 429 na 17a requisicao a 0,35s. Ritmo alvo: 4,0s.
* Nao existe filtro por intervalo de data — o corte de 2005 e aplicado depois,
  sobre o release_date do appdetails.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from .config import Filters, Settings
from .logging_setup import (fmt_elapsed, get_logger, utc_now_iso)
from .storage import append_ledger, read_json, write_json_atomic
from .tags import Taxonomy, extract_tagids_from_search_row

log = get_logger("discovery")

_ROW_RE = re.compile(r'<a[^>]*class="[^"]*search_result_row[^"]*".*?(?=<a[^>]*'
                     r'class="[^"]*search_result_row|$)', re.DOTALL)
_APPID_RE = re.compile(r'data-ds-appid="(\d+)"')
_NAME_RE = re.compile(r'<span class="title">([^<]*)</span>')
_RELEASED_RE = re.compile(r'<div class="col search_released[^"]*">\s*([^<]*?)\s*</div>')


@dataclass
class Candidate:
    app_id: int
    name: str | None = None
    search_released_raw: str | None = None
    search_tag_ids: list[int] = field(default_factory=list)
    discovered_via_tags: list[str] = field(default_factory=list)

    def merge(self, other: Candidate) -> None:
        """Une descobertas do mesmo app vindas de tags diferentes."""
        self.name = self.name or other.name
        self.search_released_raw = self.search_released_raw or other.search_released_raw
        if other.search_tag_ids and not self.search_tag_ids:
            self.search_tag_ids = other.search_tag_ids
        for tag in other.discovered_via_tags:
            if tag not in self.discovered_via_tags:
                self.discovered_via_tags.append(tag)

    def to_dict(self) -> dict[str, Any]:
        return {
            "app_id": self.app_id,
            "name": self.name,
            "search_released_raw": self.search_released_raw,
            "search_tag_ids": self.search_tag_ids,
            "discovered_via_tags": self.discovered_via_tags,
        }


def _parse_rows(results_html: str) -> list[Candidate]:
    """Extrai candidatos das linhas do results_html."""
    out: list[Candidate] = []
    for match in _ROW_RE.finditer(results_html):
        row = match.group(0)
        m_id = _APPID_RE.search(row)
        if not m_id:
            continue  # bundles/pacotes nao tem data-ds-appid unico
        m_name = _NAME_RE.search(row)
        m_rel = _RELEASED_RE.search(row)
        out.append(Candidate(
            app_id=int(m_id.group(1)),
            name=m_name.group(1).strip() if m_name else None,
            search_released_raw=m_rel.group(1).strip() if m_rel else None,
            search_tag_ids=extract_tagids_from_search_row(row),
        ))
    return out


class Discovery:
    """Enumera candidatos, uma query por tag positiva, com estado retomavel."""

    def __init__(self, client, settings: Settings, filters: Filters,  # noqa: ANN001
                 taxonomy: Taxonomy, *, pilot_pages: int | None = None) -> None:
        self.client = client
        self.settings = settings
        self.filters = filters
        self.taxonomy = taxonomy
        # Modo piloto: limita paginas por tag e grava em arquivos SEPARADOS,
        # para nao contaminar o estado nem os candidatos da coleta definitiva.
        self.pilot_pages = pilot_pages
        suffix = "_pilot" if pilot_pages else ""
        self.state_path = (f"{settings.path('checkpoints')}/"
                           f"discovery_state{suffix}.json")
        self.candidates_path = (f"{settings.path('processed')}/"
                                f"candidates{suffix}.json")
        self.indie_ledger_path = (
            f"{settings.path('processed')}/ledger_excluded_indie.json")

    # -- estado ------------------------------------------------------------
    def _load_state(self) -> dict[str, Any]:
        return read_json(self.state_path, default=None) or {
            "meta": {"started_at": utc_now_iso(),
                     "filter_version": self.filters.filter_version},
            "tags": {},
        }

    def _save_state(self, state: dict[str, Any]) -> None:
        state["meta"]["last_updated"] = utc_now_iso()
        write_json_atomic(self.state_path, state)

    # -- paginacao ---------------------------------------------------------
    def _search_page(self, tag_id: int, start: int, *,
                     untag_ids: list[int]) -> tuple[list[Candidate], int]:
        d = self.settings.discovery
        params: dict[str, Any] = {
            "query": "",
            "start": start,
            "count": d.get("page_size", 50),
            "infinite": 1,
            "json": 1,
            "category1": d.get("category1", 998),
            "os": d.get("os", "win"),
            "tags": tag_id,
        }
        # sort_by so e enviado se explicitamente configurado. Ver a nota em
        # settings.yaml: Released_DESC/Reviews_DESC FILTRAM a populacao.
        if d.get("sort_by"):
            params["sort_by"] = d["sort_by"]
        if untag_ids:
            params["untags"] = ",".join(str(t) for t in untag_ids)
        payload = self.client.get_json("search", self.settings.endpoints["search"],
                                      params)
        total = int(payload.get("total_count") or 0)
        return _parse_rows(payload.get("results_html") or ""), total

    def _enumerate_tag(self, tag_name: str, tag_id: int, *,
                       untag_ids: list[int], collector: dict[int, Candidate],
                       state: dict[str, Any], label: str) -> int:
        d = self.settings.discovery
        page_size = int(d.get("page_size", 50))
        max_pages = int(d.get("max_pages_per_tag", 2000))
        if self.pilot_pages:
            max_pages = min(max_pages, int(self.pilot_pages))

        tag_state = state["tags"].setdefault(str(tag_id), {})
        if tag_state.get("done"):
            log.info("[DISCOVERY] %-22s | JA CONCLUIDA (%s app_ids) | retomando",
                     tag_name, tag_state.get("collected", "?"))
            return 0

        start = int(tag_state.get("next_start", 0))
        seen_before = len(collector)
        total = int(tag_state.get("total_count", 0))
        # Auditoria de cobertura: sem sort_by a ordenacao do endpoint pode ser
        # instavel entre paginas, o que arriscaria nunca ver alguns itens.
        # Contabilizamos os app_ids distintos vistos NESTA tag para comparar
        # com o total_count declarado pela Steam.
        seen_this_tag: set[int] = set()

        while True:
            rows, total = self._search_page(tag_id, start, untag_ids=untag_ids)
            if start == 0:
                tag_state["total_count"] = total
                log.info("[DISCOVERY] tag %s (%d) %s -> total_count=%s",
                         tag_name, tag_id, label, f"{total:,}")
                if total == 0:
                    # Ex.: "Rail Shooter" (3954) existe na taxonomia mas nao
                    # retorna nenhum jogo (medido na FASE 1).
                    log.warning("[DISCOVERY] tag %s (%d) nao retornou nenhum jogo; "
                                "nao contribui com candidatos", tag_name, tag_id)
            if not rows:
                break

            for cand in rows:
                cand.discovered_via_tags = [tag_name]
                seen_this_tag.add(cand.app_id)
                if cand.app_id in collector:
                    collector[cand.app_id].merge(cand)
                else:
                    collector[cand.app_id] = cand

            start += page_size
            page_no = start // page_size
            total_pages = max(1, -(-total // page_size)) if total else 1
            tag_state["next_start"] = start
            tag_state["collected"] = len(collector) - seen_before

            if page_no % 5 == 0 or start >= total:
                pct = min(100.0, 100.0 * start / total) if total else 100.0
                log.info("[DISCOVERY] %-22s | pag %04d/%04d | %s app_ids | %.1f%% "
                         "| decorrido %s", tag_name, page_no, total_pages,
                         f"{len(collector):,}", pct, fmt_elapsed())
                self._save_state(state)

            if start >= total or page_no >= max_pages:
                break

        # No modo piloto a tag NAO e marcada como concluida: o piloto coleta
        # apenas uma amostra e nao deve fazer o discovery definitivo acreditar
        # que aquela tag ja foi integralmente enumerada.
        tag_state["done"] = not self.pilot_pages
        tag_state["collected"] = len(collector) - seen_before
        tag_state["distinct_seen"] = len(seen_this_tag)
        tag_state["pages_fetched"] = start // page_size

        # Cobertura: distintos vistos / total_count declarado.
        if total and not self.pilot_pages:
            cobertura = 100.0 * len(seen_this_tag) / total
            tag_state["coverage_pct"] = round(cobertura, 2)
            nivel = log.info if cobertura >= 95.0 else log.warning
            nivel("[DISCOVERY] %-22s | cobertura %.1f%% (%s distintos de %s "
                  "declarados)", tag_name, cobertura,
                  f"{len(seen_this_tag):,}", f"{total:,}")
            if cobertura < 95.0:
                log.warning("[DISCOVERY] cobertura abaixo de 95%%: a ordenacao "
                            "do endpoint pode ter sido instavel entre paginas. "
                            "Reexecutar `discover --no-resume` tende a recuperar "
                            "itens faltantes, pois a uniao e cumulativa.")
        self._save_state(state)
        return len(collector) - seen_before

    # -- execucao ----------------------------------------------------------
    def run(self, *, resume: bool = True, tag_limit: int | None = None,
            build_indie_ledger: bool = True) -> dict[str, Any]:
        """Executa o discovery e grava candidates.json + ledger de Indie."""
        positive = self.filters.positive_strong + self.filters.positive_secondary
        resolved, _ = self.taxonomy.resolve_all(positive, strict=True)
        untags, _ = self.taxonomy.resolve_all(self.filters.exclude_at_source,
                                             strict=True)
        untag_ids = sorted(untags.values())
        untag_label = ("untags=" + ",".join(f"{n}({i})" for n, i in untags.items())
                       if untags else "sem untags")

        state = self._load_state() if resume else {
            "meta": {"started_at": utc_now_iso(),
                     "filter_version": self.filters.filter_version},
            "tags": {},
        }

        collector: dict[int, Candidate] = {}
        if resume:
            for rec in (read_json(self.candidates_path, default=None) or {}
                        ).get("candidates", []):
                cand = Candidate(
                    app_id=int(rec["app_id"]), name=rec.get("name"),
                    search_released_raw=rec.get("search_released_raw"),
                    search_tag_ids=rec.get("search_tag_ids") or [],
                    discovered_via_tags=rec.get("discovered_via_tags") or [])
                collector[cand.app_id] = cand
            if collector:
                log.info("[DISCOVERY] retomando com %s candidatos ja conhecidos",
                         f"{len(collector):,}")

        tag_items = list(resolved.items())
        if tag_limit:
            tag_items = tag_items[:tag_limit]

        log.info("[DISCOVERY] %d tags positivas a consultar | %s | ritmo %.1fs/req",
                 len(tag_items), untag_label,
                 self.settings.rate_limits["search"]["min_interval_s"])

        for idx, (tag_name, tag_id) in enumerate(tag_items, start=1):
            log.info("[DISCOVERY] --- tag %d/%d: %s ---", idx, len(tag_items),
                     tag_name)
            added = self._enumerate_tag(tag_name, tag_id, untag_ids=untag_ids,
                                       collector=collector, state=state,
                                       label=untag_label)
            log.info("[DISCOVERY] %-22s | +%s novos | uniao=%s",
                     tag_name, f"{added:,}", f"{len(collector):,}")

        payload = {
            "meta": {
                "filter_version": self.filters.filter_version,
                "generated_at": utc_now_iso(),
                "discovery_method": "store_search_results",
                "category1": self.settings.discovery.get("category1"),
                "os": self.settings.discovery.get("os"),
                "positive_tags_queried": {n: i for n, i in tag_items},
                "excluded_at_source": untags,
                "total_candidates": len(collector),
                "known_bias": "storefront-only: jogos deslistados nao aparecem "
                              "(viés de sobrevivência correlacionado ao tempo)",
            },
            "candidates": [collector[k].to_dict() for k in sorted(collector)],
        }
        write_json_atomic(self.candidates_path, payload, indent=1)
        log.info("[DISCOVERY] uniao de %d tags positivas -> %s app_ids unicos",
                 len(tag_items), f"{len(collector):,}")
        log.info("[DISCOVERY] gravado em %s", self.candidates_path)

        if build_indie_ledger and untag_ids:
            self._build_indie_ledger(resolved, untag_ids, tag_items)

        return payload

    def _build_indie_ledger(self, resolved: dict[str, int],
                            untag_ids: list[int],
                            tag_items: list[tuple[str, int]]) -> None:
        """Grava os app_ids excluidos por Indie (item 2.17 / mitigacao de D2).

        Custa apenas requisicoes de discovery — sem appdetails. Torna a decisao
        Q2 REVERSIVEL: revisar o item 2.10 exigira apenas o estagio de scraping
        sobre este ledger, sem repetir o discovery.
        """
        log.info("[LEDGER]    coletando app_ids excluidos por Indie "
                 "(interseccao das tags positivas com Indie)")
        collector: dict[int, Candidate] = {}
        state = {"meta": {"started_at": utc_now_iso()}, "tags": {}}
        for tag_name, tag_id in tag_items:
            try:
                # Interseccao: tags=<positiva>,<indie> (search e AND).
                page_size = int(self.settings.discovery.get("page_size", 50))
                start, total = 0, None
                while True:
                    params = {
                        "query": "", "start": start, "count": page_size,
                        "infinite": 1, "json": 1,
                        "category1": self.settings.discovery.get("category1", 998),
                        "os": self.settings.discovery.get("os", "win"),
                        "tags": ",".join(str(t) for t in [tag_id, *untag_ids]),
                    }
                    payload = self.client.get_json(
                        "search", self.settings.endpoints["search"], params)
                    total = int(payload.get("total_count") or 0)
                    rows = _parse_rows(payload.get("results_html") or "")
                    if not rows:
                        break
                    for cand in rows:
                        cand.discovered_via_tags = [tag_name]
                        if cand.app_id in collector:
                            collector[cand.app_id].merge(cand)
                        else:
                            collector[cand.app_id] = cand
                    start += page_size
                    if start >= total:
                        break
                log.info("[LEDGER]    %-22s | Indie -> %s | acumulado=%s",
                         tag_name, f"{total:,}", f"{len(collector):,}")
            except Exception as exc:  # noqa: BLE001 - ledger nao pode derrubar o run
                log.warning("[LEDGER]    falha na tag %s: %s. O ledger fica "
                            "incompleto, mas o discovery principal e valido.",
                            tag_name, exc)
        records = [{**c.to_dict(), "exclusion_reason": "INDIE"}
                   for c in collector.values()]
        added = append_ledger(self.indie_ledger_path, records,
                              filter_version=self.filters.filter_version)
        log.info("[LEDGER]    ledger de excluidos por Indie -> %s app_ids "
                 "(+%s novos) | %s", f"{len(collector):,}", f"{added:,}",
                 self.indie_ledger_path)
        _ = state


def load_candidate_ids(processed_dir: str, *, pilot: bool = False) -> list[int]:
    suffix = "_pilot" if pilot else ""
    payload = read_json(f"{processed_dir}/candidates{suffix}.json",
                        default=None) or {}
    return [int(c["app_id"]) for c in payload.get("candidates", [])]
