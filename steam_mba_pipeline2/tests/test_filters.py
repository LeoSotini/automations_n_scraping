"""Testes dos filtros metodologicos (item 13).

Cada teste referencia o item do projeto que verifica.
"""
from __future__ import annotations

import pytest

from steamreq.config import Filters
from steamreq.filters import (ExclusionReason, InclusionBasis, build_record,
                              count_supported_languages, evaluate_tags,
                              parse_release_date)

from .fixtures import (CYBERPUNK_MINIMUM, CYBERPUNK_RECOMMENDED, appdetails,
                       tags_payload)


@pytest.fixture
def filters() -> Filters:
    return Filters.load()


def make(app_id: int, filters: Filters, *, tags: list[str] | None = None,
         tags_source: str = "STORE_HTML", **kwargs):
    payload = appdetails(app_id, **kwargs)
    tp = tags_payload(app_id, tags, tags_source)
    return build_record(app_id, payload, tp, filters, scraper_version="test")


# --- 2.3 data de lancamento ------------------------------------------------

class TestParsingDeData:
    @pytest.mark.parametrize("raw,iso,precisa", [
        ("Dec 9, 2020", "2020-12-09", True),
        ("Nov 16, 2004", "2004-11-16", True),
        ("December 9, 2020", "2020-12-09", True),
        ("9 Dec, 2020", "2020-12-09", True),
        ("Feb 24, 2022", "2022-02-24", True),
        ("December 2020", "2020-12-01", False),
        ("2020", "2020-01-01", False),
        ("Q1 2021", "2021-01-01", False),
        ("Q4 2021", "2021-10-01", False),
    ])
    def test_formatos_reconhecidos(self, raw, iso, precisa):
        d, p = parse_release_date(raw)
        assert d is not None and d.isoformat() == iso
        assert p is precisa

    @pytest.mark.parametrize("raw", ["", None, "Coming soon", "To be announced",
                                     "Feb 30, 2020"])
    def test_formatos_nao_utilizaveis(self, raw):
        d, _ = parse_release_date(raw)
        assert d is None


class TestFiltroDeData:
    def test_jogo_de_2004_e_excluido(self, filters):
        """Half-Life 2 real: release_date.date = "Nov 16, 2004"."""
        r = make(220, filters, name="Half-Life 2", release="Nov 16, 2004",
                 tags=["3D", "FPS"], minimum=CYBERPUNK_MINIMUM)
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.BEFORE_MIN_DATE
        # 2.3: a data ORIGINAL deve sobreviver ao filtro
        assert r["release_date_raw"] == "Nov 16, 2004"
        assert r["release_date"] == "2004-11-16"

    def test_jogo_de_2005_e_incluido(self, filters):
        r = make(1, filters, name="Limite", release="Jan 1, 2005", tags=["3D"])
        assert r["included_initially"] is True

    def test_jogo_recente_e_incluido(self, filters):
        r = make(2, filters, name="Recente", release="Feb 24, 2022",
                 tags=["3D", "Souls-like"])
        assert r["included_initially"] is True
        assert r["release_year"] == 2022

    def test_sem_data_utilizavel_e_preservado_para_revisao(self, filters):
        """Item 2.16: recall sobre precisao."""
        r = make(3, filters, name="Sem data", release="", tags=["3D"])
        assert r["included_initially"] is True
        assert r["needs_manual_review"] is True


# --- 2.1 / 2.2 / 2.4 ------------------------------------------------------

class TestFiltrosEstruturais:
    def test_dlc_e_excluido_mesmo_com_requisitos(self, filters):
        """DLC real (Phantom Liberty) tem pc_requirements do jogo-base."""
        r = make(2138330, filters, name="Phantom Liberty", app_type="dlc",
                 tags=["3D", "FPS"], minimum=CYBERPUNK_MINIMUM,
                 recommended=CYBERPUNK_RECOMMENDED)
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.NOT_GAME

    def test_sem_windows_e_excluido(self, filters):
        r = make(4, filters, name="Mac only", windows=False, mac=True,
                 tags=["3D"])
        assert r["exclusion_reason"] == ExclusionReason.NO_WINDOWS_SUPPORT

    def test_coming_soon_e_excluido(self, filters):
        r = make(5, filters, name="Futuro", coming_soon=True, tags=["3D"])
        assert r["exclusion_reason"] == ExclusionReason.UNRELEASED

    def test_success_false_gera_no_data(self, filters):
        r = build_record(999999999, appdetails(999999999, name="x", success=False),
                         None, filters, scraper_version="test")
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.NO_DATA
        assert r["collection_status"] == "FAILED"

    def test_ordem_de_avaliacao_do_item_2_18(self, filters):
        """Falha em varios criterios -> motivo e o PRIMEIRO da ordem."""
        r = make(6, filters, name="Multi", app_type="dlc", windows=False,
                 release="Nov 16, 2004", tags=["2D"])
        assert r["exclusion_reason"] == ExclusionReason.NOT_GAME


# --- 2.5 a 2.11 tags ------------------------------------------------------

class TestTagsPositivas:
    def test_tag_3d_inclui(self, filters):
        r = make(10, filters, name="Jogo 3D", tags=["3D", "Action"])
        assert r["included_initially"] is True
        assert r["inclusion_basis"] == InclusionBasis.STRONG_3D_TAG
        assert "3D" in r["matched_positive_tags"]

    def test_tag_forte_sem_3d_literal_inclui(self, filters):
        """Item 2.5: nao assumir que so 'literalmente 3D' pertence."""
        r = make(11, filters, name="FPS", tags=["FPS", "First-Person"])
        assert r["included_initially"] is True
        assert r["inclusion_basis"] == InclusionBasis.STRONG_3D_TAG

    def test_matched_positive_tags_registra_a_causa(self, filters):
        """Item 2.6: preservar quais tags produziram a inclusao."""
        r = make(12, filters, name="X", tags=["FPS", "Open World", "Nudity"])
        assert set(r["matched_positive_tags"]) == {"FPS", "Open World"}
        assert r["matched_positive_strong"] == ["FPS"]
        assert r["matched_positive_secondary"] == ["Open World"]

    def test_secundaria_isolada_inclui_com_revisao(self, filters):
        """Item 2.7: nao e condicao suficiente, mas 2.16 manda preservar."""
        r = make(13, filters, name="Open World 3D?", tags=["Open World"])
        assert r["included_initially"] is True
        assert r["inclusion_basis"] == InclusionBasis.SECONDARY_ONLY
        assert r["needs_manual_review"] is True

    def test_secundaria_com_2d_e_excluida(self, filters):
        """Decisao D10: sem tag forte de 3D, a evidencia negativa prevalece.

        Casos reais que entravam indevidamente antes desta regra: Terraria
        (2D + Pixel Graphics + Open World) e Hollow Knight (2D + Souls-like).
        """
        r = make(15, filters, name="Terraria", tags=["2D", "Pixel Graphics",
                                                     "Open World"])
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.TWO_DIMENSIONAL

    def test_secundaria_com_2d_platformer_e_excluida(self, filters):
        r = make(16, filters, name="Hollow Knight",
                 tags=["2D", "Souls-like", "Metroidvania"])
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.TWO_DIMENSIONAL

    def test_tag_forte_com_2d_continua_preservada(self, filters):
        """D10 NAO altera o item 2.8: com tag forte, o conflito e preservado."""
        r = make(17, filters, name="Hibrido", tags=["3D", "2D", "Open World"])
        assert r["included_initially"] is True
        assert r["tag_conflict_3d_2d"] is True
        assert r["needs_manual_review"] is True

    def test_sem_evidencia_positiva_e_excluido(self, filters):
        r = make(14, filters, name="Nada", tags=["Puzzle", "Relaxing"])
        assert r["exclusion_reason"] == ExclusionReason.NO_3D_EVIDENCE


class TestConflitos2D3D:
    """Item 2.8: 1.915 jogos reais tem AMBAS as tags. Nao descartar."""

    def test_conflito_e_preservado(self, filters):
        r = make(20, filters, name="Hibrido", tags=["3D", "2D"])
        assert r["included_initially"] is True
        assert r["tag_conflict_3d_2d"] is True
        assert r["needs_manual_review"] is True
        assert r["exclusion_reason"] is None

    def test_2d_sem_positiva_e_excluido(self, filters):
        r = make(21, filters, name="2D puro", tags=["2D", "2D Platformer"])
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.TWO_DIMENSIONAL

    def test_2_5d_nunca_exclui(self, filters):
        """Item 2.9: 2.5D pode usar assets e renderizacao 3D."""
        r = make(22, filters, name="2.5D", tags=["2.5D", "Action-Adventure"])
        assert r["included_initially"] is True
        assert "2.5D" in r["neutral_tags"]

    def test_2_5d_com_3d_nao_gera_conflito(self, filters):
        r = make(23, filters, name="X", tags=["3D", "2.5D"])
        assert r["included_initially"] is True
        assert r["tag_conflict_3d_2d"] is False


class TestTagsNegativasCandidatas:
    """Item 2.11: apenas MARCAR. 2.389 jogos sao 3D e Pixel Graphics."""

    def test_pixel_graphics_com_3d_nao_exclui(self, filters):
        r = make(30, filters, name="3D pixelado", tags=["3D", "Pixel Graphics"])
        assert r["included_initially"] is True
        assert "Pixel Graphics" in r["negative_candidate_tags"]

    def test_side_scroller_com_3d_nao_exclui(self, filters):
        r = make(31, filters, name="X", tags=["3D", "Side Scroller"])
        assert r["included_initially"] is True
        assert "Side Scroller" in r["negative_candidate_tags"]

    def test_visual_novel_sozinha_nao_gera_exclusao_propria(self, filters):
        """E excluida por AUSENCIA de evidencia 3D, nao pela tag negativa."""
        r = make(32, filters, name="VN", tags=["Visual Novel"])
        assert r["exclusion_reason"] == ExclusionReason.NO_3D_EVIDENCE


class TestTagsQueNuncaExcluem:
    """Item 2.12."""

    @pytest.mark.parametrize("tag", ["Casual", "Free to Play", "Early Access"])
    def test_nao_excluem(self, filters, tag):
        r = make(40, filters, name="X", tags=["3D", tag])
        assert r["included_initially"] is True

    def test_casual_sozinha_nao_e_motivo_de_exclusao(self, filters):
        r = make(41, filters, name="X", tags=["Casual"])
        assert r["exclusion_reason"] != "CASUAL"


class TestIndie:
    """Item 2.10 / decisao D2: excluida na FONTE, mas sempre identificada."""

    def test_indie_e_marcado(self, filters):
        r = make(50, filters, name="Indie 3D", tags=["3D", "Indie"])
        assert r["is_indie"] is True

    def test_indie_nao_e_excluido_no_estagio_de_filtro(self, filters):
        """A exclusao ocorre no discovery (untags), nao aqui. Se um Indie
        chegar ao filtro, e preservado — a validacao emite aviso."""
        r = make(51, filters, name="Indie 3D", tags=["3D", "Indie"])
        assert r["included_initially"] is True


class TestTagsIndisponiveis:
    """Age gate por login: sem base para excluir por tag (item 2.16)."""

    def test_age_gate_preserva_com_revisao(self, filters):
        r = make(60, filters, name="HuniePop", tags=None,
                 tags_source="AGE_GATE_LOGIN_REQUIRED",
                 minimum=CYBERPUNK_MINIMUM)
        assert r["included_initially"] is True
        assert r["inclusion_basis"] == InclusionBasis.TAGS_UNAVAILABLE
        assert r["needs_manual_review"] is True
        assert r["collection_status"] == "PARTIAL_NO_TAGS"
        # o dado principal e preservado
        assert r["pc_requirements"]["minimum"]["raw"] is not None


class TestAvaliacaoDeTags:
    def test_comparacao_e_insensivel_a_caixa(self, filters):
        ev = evaluate_tags(["3d", "FIRST-PERSON"], filters)
        assert ev.has_strong

    def test_tags_none_marca_indisponivel(self, filters):
        ev = evaluate_tags(None, filters)
        assert ev.tags_available is False
        assert ev.matched_positive_tags == []


# --- 2.14 threshold de reviews -------------------------------------------

class TestThresholdDeReviews:
    def test_nao_aplicado_por_padrao(self, filters):
        assert filters.review_min_threshold is None
        r = make(70, filters, name="Obscuro", tags=["3D"], reviews=3)
        assert r["included_initially"] is True

    def test_aplicado_quando_configurado(self, filters):
        filters.review_min_threshold = 1000
        r = make(71, filters, name="Obscuro", tags=["3D"], reviews=3)
        assert r["included_initially"] is False
        assert r["exclusion_reason"] == ExclusionReason.BELOW_REVIEW_THRESHOLD

    def test_metrica_oficial_e_recommendations_total(self, filters):
        r = make(72, filters, name="X", tags=["3D"], reviews=886203)
        assert r["review_count"] == 886203
        assert r["review_count_metric"] == "recommendations_total"


# --- schema e metadados ---------------------------------------------------

class TestSchema:
    def test_registro_tem_chaves_obrigatorias(self, filters):
        r = make(80, filters, name="X", tags=["3D"],
                 minimum=CYBERPUNK_MINIMUM, recommended=CYBERPUNK_RECOMMENDED)
        for key in ("app_id", "name", "steam_url", "type", "release_date_raw",
                    "release_date", "release_year", "pc_requirements",
                    "has_recommended_requirements", "included_initially",
                    "exclusion_reason", "filter_version", "scraper_version",
                    "collected_at", "source", "collection_status"):
            assert key in r

    def test_versionamento_registrado(self, filters):
        r = make(81, filters, name="X", tags=["3D"])
        assert r["filter_version"] == filters.filter_version
        assert r["scraper_version"] == "test"

    def test_url_da_loja(self, filters):
        r = make(1091500, filters, name="X", tags=["3D"])
        assert r["steam_url"] == "https://store.steampowered.com/app/1091500/"

    def test_contagem_de_idiomas(self):
        assert count_supported_languages(
            "English<strong>*</strong>, French, German") == 3
        assert count_supported_languages(None) is None
