"""Testes do cliente HTTP: backoff, retry e classificacao de erros (item 13).

Usa duplos de teste em vez de rede: os testes precisam ser rapidos e nao devem
consumir o rate limit real da Steam.
"""
from __future__ import annotations

import pytest
import requests

from steamreq.config import Settings
from steamreq.steam_client import (AgeGateLoginRequired, ForbiddenError,
                                   InvalidJSONError, NotFoundError,
                                   RateLimitedError, ServerError, SteamClient,
                                   TimeoutErrorSteam)


class FakeResponse:
    def __init__(self, status_code=200, json_data=None, text="", url="http://x",
                 headers=None):
        self.status_code = status_code
        self._json = json_data
        self.text = text
        self.url = url
        self.headers = headers or {}
        self.content = text.encode()

    def json(self):
        if self._json is None:
            raise ValueError("No JSON object could be decoded")
        return self._json


class FakeSession:
    """Devolve respostas roteirizadas e registra as chamadas."""

    def __init__(self, script):
        self.script = list(script)
        self.calls = 0
        self.cookies_seen: list[dict | None] = []
        self.headers = {}

    def get(self, url, params=None, timeout=None, cookies=None,
            allow_redirects=True):
        self.calls += 1
        self.cookies_seen.append(cookies)
        item = self.script.pop(0) if self.script else FakeResponse()
        if isinstance(item, Exception):
            raise item
        return item

    def close(self):
        pass


@pytest.fixture
def client(monkeypatch):
    settings = Settings.load()
    # Acelera os testes: sem esperas reais.
    for name in ("search", "appdetails", "store_html", "api"):
        settings.rate_limits[name]["min_interval_s"] = 0.0
        settings.rate_limits[name]["backoff_base_s"] = 0.0
    c = SteamClient(settings)
    monkeypatch.setattr("time.sleep", lambda *_: None)
    return c


class TestClassificacaoDeErros:
    def test_404_nao_faz_retry(self, client):
        client.session = FakeSession([FakeResponse(404)] * 3)
        with pytest.raises(NotFoundError):
            client.get_json("appdetails", "http://x")
        assert client.session.calls == 1, "404 e permanente: nao deve retentar"

    def test_403_nao_faz_retry(self, client):
        client.session = FakeSession([FakeResponse(403)] * 3)
        with pytest.raises(ForbiddenError):
            client.get_json("api", "http://x")
        assert client.session.calls == 1

    def test_429_faz_retry_e_depois_sucede(self, client):
        client.session = FakeSession([
            FakeResponse(429), FakeResponse(429),
            FakeResponse(200, json_data={"ok": True})])
        assert client.get_json("search", "http://x") == {"ok": True}
        assert client.session.calls == 3
        assert client.stats["rate_limited"] == 2
        assert client.stats["retries"] == 2

    def test_429_persistente_esgota_tentativas(self, client):
        client.session = FakeSession([FakeResponse(429)] * 10)
        with pytest.raises(RateLimitedError):
            client.get_json("appdetails", "http://x")
        assert client.session.calls == \
            client.settings.rate_limits["appdetails"]["max_retries"]

    def test_5xx_faz_retry(self, client):
        client.session = FakeSession([
            FakeResponse(500), FakeResponse(200, json_data={"ok": 1})])
        assert client.get_json("appdetails", "http://x") == {"ok": 1}

    def test_5xx_persistente_falha(self, client):
        client.session = FakeSession([FakeResponse(503)] * 5)
        with pytest.raises(ServerError):
            client.get_json("appdetails", "http://x")

    def test_timeout_faz_retry(self, client):
        client.session = FakeSession([
            requests.Timeout("timeout"), FakeResponse(200, json_data={"ok": 1})])
        assert client.get_json("appdetails", "http://x") == {"ok": 1}

    def test_timeout_persistente_falha(self, client):
        client.session = FakeSession([requests.Timeout("t")] * 5)
        with pytest.raises(TimeoutErrorSteam):
            client.get_json("appdetails", "http://x")

    def test_connection_error_faz_retry(self, client):
        client.session = FakeSession([
            requests.ConnectionError("down"),
            FakeResponse(200, json_data={"ok": 1})])
        assert client.get_json("appdetails", "http://x") == {"ok": 1}

    def test_json_invalido_faz_retry(self, client):
        client.session = FakeSession([
            FakeResponse(200, json_data=None, text="<html>"),
            FakeResponse(200, json_data={"ok": 1})])
        assert client.get_json("appdetails", "http://x") == {"ok": 1}

    def test_json_invalido_persistente_falha(self, client):
        client.session = FakeSession([
            FakeResponse(200, json_data=None, text="<html>")] * 5)
        with pytest.raises(InvalidJSONError):
            client.get_json("appdetails", "http://x")


class TestAgeGate:
    def test_redirect_para_login_e_detectado(self, client):
        client.session = FakeSession([FakeResponse(
            200, text="<html>login</html>",
            url="https://store.steampowered.com/login/?redir=agecheck%2Fapp%2F339800%2F")])
        with pytest.raises(AgeGateLoginRequired):
            client.get_text("store_html", "http://x", detect_age_gate=True)

    def test_pagina_normal_nao_dispara_age_gate(self, client):
        client.session = FakeSession([FakeResponse(
            200, text="<html>ok</html>",
            url="https://store.steampowered.com/app/271590/")])
        text, url = client.get_text("store_html", "http://x",
                                    detect_age_gate=True)
        assert "ok" in text and "271590" in url


class TestInterstitialDeIdade:
    """Variante (a): HTTP 200, ~51 KB, agecheck no corpo, SEM redirect.

    Foi exatamente este caso que o prototipo da FASE 3 rotulou erradamente
    como "possivel mudanca de estrutura".
    """

    def test_interstitial_e_reconhecido(self):
        from steamreq.scraper import _is_age_gate_interstitial
        html = '<html><div class="agecheck">Please enter your birth date</div></html>'
        assert _is_age_gate_interstitial(html) is True

    def test_pagina_completa_nao_e_interstitial(self):
        from steamreq.scraper import _is_age_gate_interstitial
        html = ('<html><div class="game_area_sys_req">reqs</div>'
                '<script>InitAppTagModal(1,[],"")</script></html>')
        assert _is_age_gate_interstitial(html) is False

    def test_pagina_sem_tags_e_sem_agecheck_nao_e_interstitial(self):
        """Este caso SIM deve sinalizar mudanca de estrutura."""
        from steamreq.scraper import _is_age_gate_interstitial
        assert _is_age_gate_interstitial("<html>estrutura nova</html>") is False

    def test_source_pos_age_gate_conta_como_sucesso(self):
        from steamreq.scraper import TAGS_OK_SOURCES
        assert "STORE_HTML_AFTER_AGE_GATE" in TAGS_OK_SOURCES
        assert "AGE_GATE_INTERSTITIAL" not in TAGS_OK_SOURCES


class TestBackoff:
    def test_cresce_exponencialmente(self, client):
        b = client._budgets["search"]
        b.backoff_base_s = 4.0
        client.jitter_fraction = 0.0
        assert client._backoff_seconds(b, 1) == pytest.approx(4.0)
        assert client._backoff_seconds(b, 2) == pytest.approx(8.0)
        assert client._backoff_seconds(b, 3) == pytest.approx(16.0)

    def test_respeita_o_teto(self, client):
        b = client._budgets["search"]
        b.backoff_base_s = 4.0
        client.jitter_fraction = 0.0
        client.max_backoff_s = 30.0
        assert client._backoff_seconds(b, 10) == pytest.approx(30.0)

    def test_jitter_varia_o_valor(self, client):
        b = client._budgets["search"]
        b.backoff_base_s = 10.0
        client.jitter_fraction = 0.25
        valores = {round(client._backoff_seconds(b, 1), 6) for _ in range(20)}
        assert len(valores) > 1, "jitter deveria evitar sincronizacao"

    def test_orcamentos_sao_independentes_por_endpoint(self, client):
        """Medicao da FASE 1: search e muito mais restrito que appdetails."""
        s = Settings.load()
        assert (s.rate_limits["search"]["min_interval_s"]
                > s.rate_limits["appdetails"]["min_interval_s"])
