"""Cliente HTTP com rate limiting, retry e backoff (itens 3 e 10 do projeto).

Decisoes fundamentadas nas medicoes da FASE 1:

* A Steam NAO envia header `Retry-After` nem `X-RateLimit-*` (verificado).
  Portanto o backoff e CEGO: base * 2**(tentativa-1), com jitter.
* `/search/results/` deu 429 na 17a requisicao a 0.35s/req; `/api/appdetails`
  aguentou 30/30 no mesmo ritmo. Os limites sao POR ENDPOINT, nao globais.
* `appdetails` responde HTTP 200 com `success:false` para app_id invalido.
  O status code e inutil como sinal de erro; quem checa `success` e o scraper.
"""
from __future__ import annotations

import random
import time
from dataclasses import dataclass
from typing import Any

import requests

from .logging_setup import get_logger

log = get_logger("client")


# --- taxonomia de erros (item 10) ------------------------------------------

class SteamClientError(RuntimeError):
    """Base. `code` alimenta o campo last_error do checkpoint."""

    code = "UNKNOWN"
    transient = False

    def __init__(self, message: str, *, code: str | None = None) -> None:
        super().__init__(message)
        if code:
            self.code = code


class TransientError(SteamClientError):
    """Vale a pena tentar de novo (timeout, conexao, 429, 5xx)."""

    transient = True


class PermanentError(SteamClientError):
    """Nao vale a pena tentar de novo (404, 403, redirect de age gate)."""

    transient = False


class TimeoutErrorSteam(TransientError):
    code = "TIMEOUT"


class ConnectionErrorSteam(TransientError):
    code = "CONNECTION_ERROR"


class RateLimitedError(TransientError):
    code = "HTTP_429"


class ServerError(TransientError):
    code = "HTTP_5XX"


class InvalidJSONError(TransientError):
    """Transitorio por 1 tentativa: pode ser resposta truncada."""

    code = "INVALID_JSON"


class NotFoundError(PermanentError):
    code = "HTTP_404"


class ForbiddenError(PermanentError):
    code = "HTTP_403"


class AgeGateLoginRequired(PermanentError):
    """Bloqueio por idade que EXIGE AUTENTICACAO.

    Variante que redireciona para /login/?redir=agecheck/... Nao sera
    contornada em nenhuma circunstancia (item 3 do projeto).
    """

    code = "AGE_GATE_LOGIN_REQUIRED"


class UnexpectedStatusError(PermanentError):
    code = "UNEXPECTED_STATUS"


@dataclass
class _Budget:
    """Controle de ritmo independente por endpoint."""

    min_interval_s: float
    max_retries: int
    backoff_base_s: float
    last_request_at: float = 0.0

    def wait_turn(self) -> None:
        delta = time.monotonic() - self.last_request_at
        if delta < self.min_interval_s:
            time.sleep(self.min_interval_s - delta)
        self.last_request_at = time.monotonic()


class SteamClient:
    """Sessao HTTP unica e reutilizada (keep-alive, item 3)."""

    def __init__(self, settings) -> None:  # noqa: ANN001 - evita import circular
        self.settings = settings
        http = settings.http
        self.timeout = (http.get("connect_timeout_s", 10),
                        http.get("read_timeout_s", 30))
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": " ".join(str(http["user_agent"]).split()),
            "Accept-Language": http.get("accept_language", "en-US,en;q=0.9"),
        })
        rl = settings.rate_limits
        self.jitter_fraction = float(rl.get("jitter_fraction", 0.25))
        self.max_backoff_s = float(rl.get("max_backoff_s", 120.0))
        self._budgets: dict[str, _Budget] = {}
        for name in ("search", "appdetails", "store_html", "api"):
            cfg = rl.get(name, {})
            self._budgets[name] = _Budget(
                min_interval_s=float(cfg.get("min_interval_s", 1.0)),
                max_retries=int(cfg.get("max_retries", 3)),
                backoff_base_s=float(cfg.get("backoff_base_s", 2.0)),
            )
        self.stats = {"requests": 0, "retries": 0, "rate_limited": 0, "errors": 0}

    # -- interno -----------------------------------------------------------
    def _backoff_seconds(self, budget: _Budget, attempt: int) -> float:
        raw = budget.backoff_base_s * (2 ** (attempt - 1))
        raw = min(raw, self.max_backoff_s)
        jitter = raw * self.jitter_fraction
        return max(0.5, raw + random.uniform(-jitter, jitter))

    @staticmethod
    def _classify(response: requests.Response) -> SteamClientError | None:
        code = response.status_code
        if code == 200:
            return None
        if code == 429:
            return RateLimitedError(f"rate limited em {response.url}")
        if 500 <= code < 600:
            return ServerError(f"HTTP {code} em {response.url}")
        if code == 404:
            return NotFoundError(f"HTTP 404 em {response.url}")
        if code == 403:
            return ForbiddenError(f"HTTP 403 em {response.url}")
        return UnexpectedStatusError(f"HTTP {code} em {response.url}")

    def _request(self, budget_name: str, url: str, *,
                 params: dict[str, Any] | None = None,
                 expect_json: bool,
                 detect_age_gate: bool = False,
                 cookies: dict[str, str] | None = None) -> tuple[Any, requests.Response]:
        budget = self._budgets[budget_name]
        last_error: SteamClientError = TransientError("nenhuma tentativa executada")

        for attempt in range(1, budget.max_retries + 1):
            budget.wait_turn()
            try:
                self.stats["requests"] += 1
                resp = self.session.get(url, params=params, timeout=self.timeout,
                                        cookies=cookies, allow_redirects=True)
            except requests.Timeout as exc:
                last_error = TimeoutErrorSteam(f"timeout em {url}: {exc}")
            except requests.ConnectionError as exc:
                last_error = ConnectionErrorSteam(f"falha de conexao em {url}: {exc}")
            except requests.RequestException as exc:  # rede/protocolo inesperado
                last_error = TransientError(f"erro de requisicao em {url}: {exc}")
            else:
                # Age gate por login: bloqueio deliberado, nao contornar.
                if detect_age_gate and "/login/" in resp.url and "agecheck" in resp.url:
                    raise AgeGateLoginRequired(
                        f"age gate exige autenticacao: {resp.url}")

                err = self._classify(resp)
                if err is None:
                    if not expect_json:
                        return resp.text, resp
                    try:
                        return resp.json(), resp
                    except ValueError as exc:
                        last_error = InvalidJSONError(
                            f"JSON invalido em {resp.url} "
                            f"({len(resp.content)} bytes): {exc}")
                elif isinstance(err, PermanentError):
                    self.stats["errors"] += 1
                    log.debug("erro permanente (%s) em %s", err.code, url)
                    raise err
                else:
                    last_error = err
                    if isinstance(err, RateLimitedError):
                        self.stats["rate_limited"] += 1

            if attempt < budget.max_retries:
                wait = self._backoff_seconds(budget, attempt)
                self.stats["retries"] += 1
                log.warning("[RATE] %s (%s) -> backoff %.1fs "
                            "(tentativa %d/%d) | %s",
                            last_error.code, budget_name, wait, attempt + 1,
                            budget.max_retries, url.split("?")[0])
                time.sleep(wait)

        self.stats["errors"] += 1
        raise last_error

    # -- API publica -------------------------------------------------------
    def get_json(self, budget: str, url: str,
                 params: dict[str, Any] | None = None) -> Any:
        data, _ = self._request(budget, url, params=params, expect_json=True)
        return data

    def get_text(self, budget: str, url: str,
                 params: dict[str, Any] | None = None, *,
                 detect_age_gate: bool = False,
                 cookies: dict[str, str] | None = None) -> tuple[str, str]:
        """Retorna (texto, url_final)."""
        text, resp = self._request(budget, url, params=params, expect_json=False,
                                   detect_age_gate=detect_age_gate,
                                   cookies=cookies)
        return text, resp.url

    def close(self) -> None:
        self.session.close()

    def __enter__(self) -> SteamClient:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()
