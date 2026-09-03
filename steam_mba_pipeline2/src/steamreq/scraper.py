"""Estagio 2 — SCRAPE. Coleta appdetails + tags por app_id (itens 7, 9, 10).

Garantias:
* Uma falha individual NUNCA interrompe o lote (item 10): cada app roda em
  try/except isolado, o erro vai para o checkpoint e o loop continua.
* Idempotente: app com raw presente e checkpoint COMPLETE e pulado SEM
  requisicao de rede. Reexecutar e seguro e barato.
* Interrupcao (Ctrl+C) faz flush do checkpoint e encerra com resumo. Retomar
  no jogo 4.731 continua no 4.732.
* HTTP 200 com `success:false` e tratado como DADO AUSENTE, nao como falha de
  rede — distincao exigida pelo item 8.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Iterable

from .config import Filters, Settings
from .logging_setup import (fmt_bytes, fmt_elapsed, fmt_eta, fmt_progress,
                            get_logger, utc_now_iso)
from .steam_client import (AgeGateLoginRequired, PermanentError,
                           SteamClientError)
from .storage import Checkpoint, RawStore, Status
from .tags import extract_tags_from_html

log = get_logger("scraper")

# Marcadores do interstitial de idade renderizado na propria pagina.
# Medido: a pagina bloqueada tem ~51 KB, contem `agecheck`/`age_gate` e NAO
# contem `game_area_sys_req` nem `InitAppTagModal`.
_AGE_GATE_MARKERS = ("agecheck", "age_gate", "agegate")

# tags_source que representam coleta bem-sucedida das tags.
TAGS_OK_SOURCES = ("STORE_HTML", "STORE_HTML_AFTER_AGE_GATE")


def _is_age_gate_interstitial(html: str) -> bool:
    low = html.lower()
    return (any(m in low for m in _AGE_GATE_MARKERS)
            and "game_area_sys_req" not in html)


@dataclass
class ScrapeStats:
    completed: int = 0
    partial: int = 0
    failed: int = 0
    skipped: int = 0
    started_at: float = field(default_factory=time.monotonic)

    @property
    def processed(self) -> int:
        return self.completed + self.partial + self.failed

    @property
    def elapsed_s(self) -> float:
        return time.monotonic() - self.started_at


class Scraper:
    def __init__(self, client, settings: Settings, filters: Filters) -> None:  # noqa: ANN001
        self.client = client
        self.settings = settings
        self.filters = filters
        self.raw = RawStore(settings.path("raw_appdetails"), settings.path("raw_tags"))
        self.checkpoint = Checkpoint(
            f"{settings.path('checkpoints')}/scrape_state.json",
            filter_version=filters.filter_version,
            scraper_version=settings.scraper_version,
            flush_every=int(settings.scrape.get("checkpoint_every", 25)),
        )
        self.fetch_tags = bool(settings.scrape.get("fetch_tags_from_html", True))
        self.max_attempts = int(settings.scrape.get("max_attempts", 3))
        self.accept_age_gate = bool(
            settings.scrape.get("accept_age_gate_interstitial", False))
        self.age_gate_cookies = {
            str(k): str(v) for k, v in
            (settings.scrape.get("age_gate_cookies") or {}).items()}
        self.sleep_between_apps_s = float(
            settings.scrape.get("sleep_between_apps_s", 0.0))

    # -- unidades de coleta ------------------------------------------------
    def _fetch_appdetails(self, app_id: int) -> tuple[dict, int]:
        payload = self.client.get_json(
            "appdetails", self.settings.endpoints["appdetails"],
            {"appids": app_id, "l": "english", "cc": "us"})
        if not isinstance(payload, dict):
            raise SteamClientError(f"appdetails devolveu {type(payload).__name__} "
                                   "em vez de objeto", code="UNEXPECTED_PAYLOAD")
        size = len(str(payload))
        return payload, size

    def _fetch_tags(self, app_id: int) -> dict:
        """Busca as 20 tags do HTML (decisao Q3/D3).

        O appdetails NAO retorna tags — verificado em 9 jogos na FASE 1.

        Trata as DUAS variantes de age gate observadas empiricamente:
          (a) interstitial na propria pagina: HTTP 200, ~51 KB, sem
              InitAppTagModal, com marcador `agecheck` no corpo. Resolvido por
              auto-declaracao de data de nascimento, SEM autenticacao. O
              tratamento e controlado por accept_age_gate_interstitial.
          (b) redirect para /login/?redir=agecheck/...: EXIGE AUTENTICACAO e
              NUNCA e contornado (item 3 do projeto).
        """
        url = self.settings.endpoints["store_page"].format(app_id=app_id)

        def fetch(with_cookies: bool) -> tuple[str, str]:
            return self.client.get_text(
                "store_html", url, detect_age_gate=True,
                cookies=self.age_gate_cookies if with_cookies else None)

        def result(tags, source, final_url) -> dict:
            return {"app_id": app_id, "tags": tags or None,
                    "tags_source": source, "collected_at": utc_now_iso(),
                    "source_url": final_url}

        try:
            html, final_url = fetch(False)
        except AgeGateLoginRequired:
            # Variante (b): bloqueio por autenticacao. Nao contornar.
            return result(None, "AGE_GATE_LOGIN_REQUIRED", url)

        tags = extract_tags_from_html(html)
        if tags:
            return result(tags, "STORE_HTML", final_url)

        # Variante (a): interstitial na propria pagina.
        if _is_age_gate_interstitial(html):
            if not self.accept_age_gate:
                log.warning("   [WARN] age gate na pagina (interstitial); "
                            "accept_age_gate_interstitial=false -> tags omitidas")
                return result(None, "AGE_GATE_INTERSTITIAL", final_url)
            log.debug("age gate na pagina para app %d; reenviando com "
                      "auto-declaracao de idade", app_id)
            try:
                html, final_url = fetch(True)
            except AgeGateLoginRequired:
                return result(None, "AGE_GATE_LOGIN_REQUIRED", url)
            tags = extract_tags_from_html(html)
            if tags:
                return result(tags, "STORE_HTML_AFTER_AGE_GATE", final_url)
            log.warning("   [WARN] age gate persistiu apos auto-declaracao "
                        "(app %d) -> tags indisponiveis", app_id)
            return result(None, "AGE_GATE_INTERSTITIAL", final_url)

        # Nem tags nem age gate: sinal genuino de mudanca de estrutura (item 10).
        log.warning("   [WARN] InitAppTagModal ausente e SEM age gate para app "
                    "%d (%s) -> possivel mudanca de estrutura da pagina",
                    app_id, fmt_bytes(len(html)))
        return result(None, "UNAVAILABLE", final_url)

    # -- processamento de um app ------------------------------------------
    def _process_one(self, app_id: int, *, force: bool) -> Status:
        cp = self.checkpoint

        if not force and cp.is_done(app_id) and self.raw.has_appdetails(app_id):
            return Status.COMPLETE if self.raw.has_tags(app_id) else Status.PARTIAL_NO_TAGS

        # --- appdetails (obrigatorio) ---
        if not force and self.raw.has_appdetails(app_id):
            payload = self.raw.load_appdetails(app_id) or {}
            size = len(str(payload))
            log.debug("appdetails de cache para app %d", app_id)
        else:
            payload, size = self._fetch_appdetails(app_id)
            self.raw.save_appdetails(app_id, payload)

        node = payload.get(str(app_id)) or {}
        if not node.get("success"):
            # DADO AUSENTE na Steam (app removido/indisponivel na regiao), nao
            # falha de rede. Sem retry — retentar nao mudaria o resultado.
            log.warning("   [FAIL] appdetails success=false (app removido ou "
                        "indisponivel na regiao) | sem retry")
            cp.mark(app_id, Status.FAILED, error="APPDETAILS_SUCCESS_FALSE",
                    increment_attempt=True, has_appdetails=True, has_tags=False)
            return Status.FAILED

        data = node.get("data") or {}
        pcr = data.get("pc_requirements")
        has_rec = isinstance(pcr, dict) and bool(pcr.get("recommended"))
        has_min = isinstance(pcr, dict) and bool(pcr.get("minimum"))
        log.info("   [OK] appdetails %s | type=%s | release=%s | reviews=%s",
                 fmt_bytes(size), data.get("type"),
                 (data.get("release_date") or {}).get("date"),
                 f"{(data.get('recommendations') or {}).get('total') or 0:,}")

        # --- tags (decisao Q3) ---
        tags_source = "SKIPPED"
        if self.fetch_tags:
            if not force and self.raw.has_tags(app_id):
                tags_payload = self.raw.load_tags(app_id) or {}
            else:
                tags_payload = self._fetch_tags(app_id)
                self.raw.save_tags(app_id, tags_payload)
            tags_source = tags_payload.get("tags_source") or "UNAVAILABLE"
            n_tags = len(tags_payload.get("tags") or [])
            if tags_source in TAGS_OK_SOURCES:
                log.info("   [OK] %d tags via %s", n_tags, tags_source)
            else:
                log.warning("   [WARN] %s -> tags indisponiveis; requisitos "
                            "preservados", tags_source)

        # --- requisitos: relatorio verbose (item 9) ---
        if has_rec:
            from .requirements_parser import detect_markup_format
            fmt = detect_markup_format(pcr.get("recommended") or "")
            log.info("   [OK] requisitos RECOMENDADOS capturados (formato %s)", fmt)
        elif has_min:
            log.info("   [INFO] sem requisitos recomendados na Steam "
                     "(has_recommended_requirements=false); minimos capturados")
        else:
            log.info("   [INFO] jogo sem requisitos de PC publicados")

        status = (Status.COMPLETE if tags_source in (*TAGS_OK_SOURCES, "SKIPPED")
                  else Status.PARTIAL_NO_TAGS)
        cp.mark(app_id, status,
                error=None if status is Status.COMPLETE else tags_source,
                increment_attempt=True, has_appdetails=True,
                has_tags=tags_source in TAGS_OK_SOURCES,
                has_recommended=has_rec, has_minimum=has_min,
                name=data.get("name"))
        return status

    # -- loop principal ----------------------------------------------------
    def run(self, app_ids: Iterable[int], *, force: bool = False) -> ScrapeStats:
        ids = list(dict.fromkeys(int(a) for a in app_ids))  # dedup preservando ordem
        total = len(ids)
        stats = ScrapeStats()
        cp = self.checkpoint
        cp.register_pending(ids)

        log.info("[SCRAPE]    %s app_ids na fila | tags do HTML=%s | "
                 "ritmo appdetails=%.1fs | pausa entre jogos=%.1fs",
                 f"{total:,}", self.fetch_tags,
                 self.settings.rate_limits["appdetails"]["min_interval_s"],
                 self.sleep_between_apps_s)
        if self.sleep_between_apps_s > 0 and total > 50:
            est_h = total * (self.sleep_between_apps_s
                             + (1.6 if self.fetch_tags else 0.6)) / 3600
            log.info("[SCRAPE]    estimativa de duracao: %.1f h (%.1f dias) "
                     "no ritmo configurado", est_h, est_h / 24)

        try:
            for idx, app_id in enumerate(ids, start=1):
                if not force and cp.is_done(app_id) and self.raw.has_appdetails(app_id):
                    stats.skipped += 1
                    log.debug("[SCRAPE %s] app_id=%d | SKIP (ja concluido)",
                              fmt_progress(idx, total), app_id)
                    continue

                attempts = cp.attempts_of(app_id)
                if attempts >= self.max_attempts and cp.status_of(app_id) is Status.FAILED:
                    stats.skipped += 1
                    log.info("[SCRAPE %s] app_id=%d | SKIP (%d tentativas "
                             "esgotadas; use retry-failed para forcar)",
                             fmt_progress(idx, total), app_id, attempts)
                    continue

                name = (cp.get(app_id) or {}).get("name") or ""
                log.info("[SCRAPE %s] app_id=%d%s", fmt_progress(idx, total),
                         app_id, f" | {name}" if name else "")

                try:
                    status = self._process_one(app_id, force=force)
                except AgeGateLoginRequired as exc:
                    stats.partial += 1
                    log.warning("   [WARN] %s", exc)
                    cp.mark(app_id, Status.PARTIAL_NO_TAGS,
                            error="AGE_GATE_LOGIN_REQUIRED", increment_attempt=True)
                except PermanentError as exc:
                    stats.failed += 1
                    log.error("   [FAIL] %s (%s) | sem retry", exc, exc.code)
                    cp.mark(app_id, Status.FAILED, error=exc.code,
                            increment_attempt=True)
                except SteamClientError as exc:
                    stats.failed += 1
                    log.error("   [FAIL] %s (%s) | tentativas de rede esgotadas",
                              exc, exc.code)
                    cp.mark(app_id, Status.FAILED, error=exc.code,
                            increment_attempt=True)
                except Exception as exc:  # noqa: BLE001 - isolamento por app (item 10)
                    stats.failed += 1
                    log.exception("   [FAIL] erro inesperado em app %d: %s",
                                  app_id, exc)
                    cp.mark(app_id, Status.FAILED,
                            error=f"UNEXPECTED:{type(exc).__name__}",
                            increment_attempt=True)
                else:
                    if status is Status.COMPLETE:
                        stats.completed += 1
                    elif status is Status.PARTIAL_NO_TAGS:
                        stats.partial += 1
                    else:
                        stats.failed += 1

                if idx % int(self.settings.scrape.get("checkpoint_every", 25)) == 0:
                    self._log_checkpoint(stats, idx, total)

                # Pausa entre jogos. Aplicada apenas depois de um app que
                # efetivamente gerou requisicoes — apps pulados por checkpoint
                # sairam do loop antes deste ponto, para que a retomada de uma
                # coleta longa nao pague o custo novamente.
                if self.sleep_between_apps_s > 0 and idx < total:
                    time.sleep(self.sleep_between_apps_s)

        except KeyboardInterrupt:
            log.warning("\n[INTERRUPT] interrupcao manual detectada; gravando "
                        "checkpoint antes de sair")
            cp.flush()
            self._log_checkpoint(stats, stats.processed + stats.skipped, total)
            log.warning("[INTERRUPT] estado preservado. Retome com "
                        "`python main.py resume`")
            return stats

        cp.flush()
        self._log_checkpoint(stats, total, total)
        return stats

    def _log_checkpoint(self, stats: ScrapeStats, done: int, total: int) -> None:
        self.checkpoint.flush()
        counts = self.checkpoint.counts()
        remaining = max(0, total - done)
        log.info("[CHECKPOINT] %s completos | %s parciais | %s falhas | "
                 "%s pulados | %s restantes",
                 f"{stats.completed:,}", f"{stats.partial:,}",
                 f"{stats.failed:,}", f"{stats.skipped:,}", f"{remaining:,}")
        log.info("             estado global: COMPLETE=%s PARTIAL=%s FAILED=%s "
                 "PENDING=%s", f"{counts.get('COMPLETE', 0):,}",
                 f"{counts.get('PARTIAL_NO_TAGS', 0):,}",
                 f"{counts.get('FAILED', 0):,}", f"{counts.get('PENDING', 0):,}")
        log.info("             decorrido %s | ETA %s | HTTP: %s reqs, %s retries, "
                 "%s rate-limited", fmt_elapsed(stats.elapsed_s),
                 fmt_eta(stats.processed, total - stats.skipped, stats.elapsed_s),
                 f"{self.client.stats['requests']:,}",
                 f"{self.client.stats['retries']:,}",
                 f"{self.client.stats['rate_limited']:,}")
