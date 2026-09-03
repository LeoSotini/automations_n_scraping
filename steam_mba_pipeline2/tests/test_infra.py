"""Testes de taxonomia de tags, storage/checkpoint e validacoes (item 13)."""
from __future__ import annotations

import json
import os

import pytest

from steamreq.config import ConfigError, Filters
from steamreq.storage import (Checkpoint, RawStore, Status, append_ledger,
                              read_json, write_json_atomic)
from steamreq.tags import (TagResolutionError, Taxonomy,
                           extract_tagids_from_search_row,
                           extract_tags_from_html)
from steamreq.validators import validate_dataset

from .fixtures import CYBERPUNK_MINIMUM, CYBERPUNK_RECOMMENDED, appdetails, tags_payload


# --- taxonomia -------------------------------------------------------------

TAX_PAYLOAD = {"response": {"tags": [
    {"tagid": 4191, "name": "3D"},
    {"tagid": 3871, "name": "2D"},
    {"tagid": 4975, "name": "2.5D"},
    {"tagid": 492, "name": "Indie"},
    {"tagid": 1663, "name": "FPS"},
    {"tagid": 4166, "name": "Atmospheric"},
]}}


class TestTaxonomia:
    def test_resolucao_por_nome(self):
        tax = Taxonomy.from_payload(TAX_PAYLOAD)
        assert tax.resolve("3D") == 4191
        assert tax.resolve("Indie") == 492

    def test_4166_e_atmospheric_nao_3d(self):
        """Guard-rail contra o erro real cometido na FASE 1."""
        tax = Taxonomy.from_payload(TAX_PAYLOAD)
        assert tax.by_id[4166] == "Atmospheric"
        assert tax.resolve("3D") != 4166

    def test_tolera_caixa_e_espacos(self):
        tax = Taxonomy.from_payload(TAX_PAYLOAD)
        assert tax.resolve("  fps ") == 1663

    def test_tag_inexistente_retorna_none(self):
        tax = Taxonomy.from_payload(TAX_PAYLOAD)
        assert tax.resolve("Homemade") is None

    def test_falha_ruidosa_em_modo_strict(self):
        """Prosseguir produziria um filtro silenciosamente vazio."""
        tax = Taxonomy.from_payload(TAX_PAYLOAD)
        with pytest.raises(TagResolutionError, match="Homemade"):
            tax.resolve_all(["3D", "Homemade"], strict=True)

    def test_modo_nao_strict_reporta_faltantes(self):
        tax = Taxonomy.from_payload(TAX_PAYLOAD)
        resolved, missing = tax.resolve_all(["3D", "Homemade"], strict=False)
        assert resolved == {"3D": 4191}
        assert missing == ["Homemade"]

    def test_payload_sem_tags_falha(self):
        with pytest.raises(TagResolutionError):
            Taxonomy.from_payload({"response": {}})

    def test_filters_yaml_real_resolve_contra_taxonomia_real(self):
        """Usa o snapshot real, se existir, para validar o filters.yaml."""
        path = os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), "data", "raw", "tag_taxonomy.json")
        if not os.path.exists(path):
            pytest.skip("snapshot da taxonomia ainda nao baixado")
        cached = read_json(path)
        tax = Taxonomy.from_payload(cached.get("payload", cached))
        filters = Filters.load()
        resolved, missing = tax.resolve_all(filters.all_referenced_tags,
                                            strict=False)
        assert not missing, f"tags do filters.yaml ausentes na taxonomia: {missing}"
        assert resolved["3D"] == 4191


class TestExtracaoDeTagsDoHtml:
    def test_init_app_tag_modal(self):
        html = ('<script>InitAppTagModal( 1091500, '
                '[{"tagid":4115,"name":"Cyberpunk"},'
                '{"tagid":1695,"name":"Open World"}], "..." );</script>')
        tags = extract_tags_from_html(html)
        assert tags == [{"tagid": 4115, "name": "Cyberpunk"},
                        {"tagid": 1695, "name": "Open World"}]

    def test_ausencia_retorna_lista_vazia(self):
        assert extract_tags_from_html("<html>nada aqui</html>") == []

    def test_json_invalido_nao_lanca(self):
        html = 'InitAppTagModal( 1, [{"tagid":1,"name":], "x" );'
        assert extract_tags_from_html(html) == []

    def test_tagids_da_linha_do_search(self):
        row = '<a data-ds-appid="123" data-ds-tagids="[3964,3798,492]">x</a>'
        assert extract_tagids_from_search_row(row) == [3964, 3798, 492]

    def test_linha_sem_tagids(self):
        assert extract_tagids_from_search_row('<a data-ds-appid="1"></a>') == []


# --- storage / checkpoint --------------------------------------------------

class TestEscritaAtomica:
    def test_grava_e_le(self, tmp_path):
        p = str(tmp_path / "x.json")
        write_json_atomic(p, {"a": 1})
        assert read_json(p) == {"a": 1}

    def test_nao_deixa_tmp_orfao(self, tmp_path):
        p = str(tmp_path / "x.json")
        write_json_atomic(p, {"a": 1})
        assert not [f for f in os.listdir(tmp_path) if f.endswith(".tmp")]

    def test_utf8_preservado(self, tmp_path):
        p = str(tmp_path / "x.json")
        write_json_atomic(p, {"nome": "Kingdom Come: Deliverance — Ação"})
        with open(p, encoding="utf-8") as fh:
            assert "Ação" in fh.read()

    def test_json_corrompido_e_tratado_como_ausente(self, tmp_path):
        p = tmp_path / "bad.json"
        p.write_text("{ isso nao e json", encoding="utf-8")
        assert read_json(str(p), default="fallback") == "fallback"

    def test_arquivo_anterior_sobrevive_a_falha(self, tmp_path):
        p = str(tmp_path / "x.json")
        write_json_atomic(p, {"bom": True})
        with pytest.raises(TypeError):
            write_json_atomic(p, {"ruim": object()})
        assert read_json(p) == {"bom": True}


class TestCheckpoint:
    def _cp(self, tmp_path) -> Checkpoint:
        return Checkpoint(str(tmp_path / "state.json"),
                          filter_version="1.0.0", scraper_version="0.1.0",
                          flush_every=1)

    def test_estado_inicial_e_pending(self, tmp_path):
        cp = self._cp(tmp_path)
        assert cp.status_of(730) is Status.PENDING
        assert cp.is_done(730) is False

    def test_marcacao_e_persistencia(self, tmp_path):
        cp = self._cp(tmp_path)
        cp.mark(730, Status.COMPLETE, increment_attempt=True)
        cp.flush()
        cp2 = self._cp(tmp_path)
        assert cp2.status_of(730) is Status.COMPLETE
        assert cp2.attempts_of(730) == 1

    def test_resume_nao_reprocessa_concluidos(self, tmp_path):
        cp = self._cp(tmp_path)
        for aid in (1, 2, 3):
            cp.mark(aid, Status.COMPLETE)
        cp.mark(4, Status.FAILED, error="HTTP_404")
        cp.flush()
        cp2 = self._cp(tmp_path)
        pendentes = [a for a in (1, 2, 3, 4, 5) if not cp2.is_done(a)]
        assert pendentes == [4, 5]

    def test_interrupcao_no_meio_preserva_progresso(self, tmp_path):
        """Simula o cenario do item 7: parar no app 4.731."""
        cp = Checkpoint(str(tmp_path / "state.json"), filter_version="1.0.0",
                        scraper_version="0.1.0", flush_every=10_000)
        for aid in range(1, 4732):
            cp.mark(aid, Status.COMPLETE)
        cp.flush()
        cp2 = self._cp(tmp_path)
        assert cp2.is_done(4731) is True
        assert cp2.is_done(4732) is False
        assert cp2.counts()["COMPLETE"] == 4731

    def test_partial_conta_como_concluido(self, tmp_path):
        cp = self._cp(tmp_path)
        cp.mark(339800, Status.PARTIAL_NO_TAGS, error="AGE_GATE_LOGIN_REQUIRED")
        assert cp.is_done(339800) is True

    def test_reset_failed(self, tmp_path):
        cp = self._cp(tmp_path)
        cp.mark(1, Status.FAILED, error="HTTP_404")
        cp.mark(2, Status.COMPLETE)
        assert cp.reset_failed() == 1
        assert cp.status_of(1) is Status.PENDING
        assert cp.status_of(2) is Status.COMPLETE

    def test_register_pending_nao_sobrescreve(self, tmp_path):
        cp = self._cp(tmp_path)
        cp.mark(1, Status.COMPLETE)
        cp.register_pending([1, 2, 3])
        assert cp.status_of(1) is Status.COMPLETE
        assert cp.status_of(2) is Status.PENDING

    def test_erro_e_timestamp_registrados(self, tmp_path):
        cp = self._cp(tmp_path)
        cp.mark(1, Status.FAILED, error="HTTP_429", increment_attempt=True)
        e = cp.get(1)
        assert e["last_error"] == "HTTP_429"
        assert e["last_attempt_at"] is not None
        assert e["attempts"] == 1

    def test_mudanca_de_filter_version_gera_aviso(self, tmp_path, caplog):
        cp = self._cp(tmp_path)
        cp.mark(1, Status.COMPLETE)
        cp.flush()
        with caplog.at_level("WARNING"):
            Checkpoint(str(tmp_path / "state.json"), filter_version="2.0.0",
                       scraper_version="0.1.0")
        assert any("filter_version" in r.message for r in caplog.records)


class TestRawStore:
    def test_idempotencia(self, tmp_path):
        raw = RawStore(str(tmp_path / "ad"), str(tmp_path / "tg"))
        assert raw.has_appdetails(730) is False
        raw.save_appdetails(730, {"730": {"success": True}})
        assert raw.has_appdetails(730) is True
        assert raw.load_appdetails(730)["730"]["success"] is True

    def test_known_app_ids(self, tmp_path):
        raw = RawStore(str(tmp_path / "ad"), str(tmp_path / "tg"))
        for aid in (730, 220, 1091500):
            raw.save_appdetails(aid, {})
        assert raw.known_app_ids() == [220, 730, 1091500]


class TestLedger:
    def test_dedup_por_app_id(self, tmp_path):
        p = str(tmp_path / "ledger.json")
        assert append_ledger(p, [{"app_id": 1}, {"app_id": 2}],
                             filter_version="1.0.0") == 2
        assert append_ledger(p, [{"app_id": 2}, {"app_id": 3}],
                             filter_version="1.0.0") == 1
        data = read_json(p)
        assert [r["app_id"] for r in data["records"]] == [1, 2, 3]
        assert data["meta"]["total"] == 3


# --- validacoes -----------------------------------------------------------

class TestValidacoes:
    @pytest.fixture
    def filters(self):
        return Filters.load()

    def _rec(self, filters, app_id=1091500, **kw):
        from steamreq.filters import build_record
        payload = appdetails(app_id, name="X", minimum=CYBERPUNK_MINIMUM,
                             recommended=CYBERPUNK_RECOMMENDED, **kw)
        return build_record(app_id, payload,
                            tags_payload(app_id, ["3D", "FPS"]), filters,
                            scraper_version="test")

    def test_registro_valido_passa(self, filters):
        rep = validate_dataset([self._rec(filters)], filters)
        assert rep.ok, [i.as_dict() for i in rep.errors]

    def test_sem_recomendados_nao_e_erro(self, filters):
        from steamreq.filters import build_record
        payload = appdetails(620, name="Portal 2", release="Apr 18, 2011",
                             minimum=CYBERPUNK_MINIMUM)
        rec = build_record(620, payload, tags_payload(620, ["3D"]), filters,
                          scraper_version="test")
        rep = validate_dataset([rec], filters)
        assert rec["has_recommended_requirements"] is False
        assert rep.ok, [i.as_dict() for i in rep.errors]

    def test_app_id_duplicado_e_erro(self, filters):
        r = self._rec(filters)
        rep = validate_dataset([r, dict(r)], filters)
        assert "DUPLICATE_APP_ID" in rep.counts_by_code()

    def test_incluido_antes_de_2005_e_erro(self, filters):
        r = self._rec(filters)
        r["release_date"] = "2003-01-01"
        r["release_year"] = 2003
        rep = validate_dataset([r], filters)
        assert "INCLUDED_BEFORE_MIN_DATE" in rep.counts_by_code()

    def test_excluido_sem_motivo_e_erro(self, filters):
        r = self._rec(filters)
        r["included_initially"] = False
        r["exclusion_reason"] = None
        rep = validate_dataset([r], filters)
        assert "EXCLUDED_WITHOUT_REASON" in rep.counts_by_code()

    def test_flag_de_recomendados_incoerente_e_erro(self, filters):
        r = self._rec(filters)
        r["has_recommended_requirements"] = False
        rep = validate_dataset([r], filters)
        assert "RECOMMENDED_FLAG_MISMATCH" in rep.counts_by_code()

    def test_ano_incoerente_com_data_e_erro(self, filters):
        r = self._rec(filters)
        r["release_year"] = 1999
        rep = validate_dataset([r], filters)
        assert "YEAR_DATE_MISMATCH" in rep.counts_by_code()

    def test_conflito_sem_flag_de_revisao_e_erro(self, filters):
        r = self._rec(filters)
        r["tag_conflict_3d_2d"] = True
        r["needs_manual_review"] = False
        rep = validate_dataset([r], filters)
        assert "CONFLICT_NOT_FLAGGED" in rep.counts_by_code()

    def test_serializavel_em_json(self, filters):
        r = self._rec(filters)
        json.dumps(r, ensure_ascii=False)


# --- configuracao ---------------------------------------------------------

class TestConfigGuardRails:
    def test_filters_yaml_real_carrega(self):
        f = Filters.load()
        assert f.filter_version
        assert "3D" in f.positive_strong
        assert "2.5D" in f.neutral_never_exclude
        assert "Homemade" not in f.all_referenced_tags

    def test_violacao_do_item_2_12_e_rejeitada(self, tmp_path):
        import yaml
        cfg = {"filter_version": "x", "tags": {
            "positive_strong": ["3D"], "exclude_at_source": ["Casual"],
            "never_exclude": ["Casual"]}}
        p = tmp_path / "f.yaml"
        p.write_text(yaml.safe_dump(cfg), encoding="utf-8")
        with pytest.raises(ConfigError, match="2.12"):
            Filters.load(str(p))

    def test_violacao_do_item_2_9_e_rejeitada(self, tmp_path):
        import yaml
        cfg = {"filter_version": "x", "tags": {
            "positive_strong": ["3D"], "negative_2d": ["2.5D"],
            "neutral_never_exclude": ["2.5D"]}}
        p = tmp_path / "f.yaml"
        p.write_text(yaml.safe_dump(cfg), encoding="utf-8")
        with pytest.raises(ConfigError, match="2.9"):
            Filters.load(str(p))
