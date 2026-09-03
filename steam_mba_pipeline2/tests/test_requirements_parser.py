"""Testes do parser de requisitos (item 13 do projeto).

O teste mais importante deste arquivo e o do FORMATO B: e o caso que um parser
ingenuo silenciosamente transforma em campos vazios.
"""
from __future__ import annotations

import pytest

from steamreq.requirements_parser import (detect_markup_format, empty_block,
                                          parse_pc_requirements,
                                          parse_requirements_block)

from .fixtures import (CYBERPUNK_MINIMUM, CYBERPUNK_RECOMMENDED,
                       HADES_RECOMMENDED, PORTAL2_MINIMUM,
                       TERRARIA_RECOMMENDED)


class TestFormatoA:
    """Formato canonico: <li><strong>Rotulo:</strong> valor."""

    def test_parsing_dos_requisitos_recomendados(self):
        b = parse_requirements_block(CYBERPUNK_RECOMMENDED)
        assert b["os"] == "64-bit Windows 10"
        assert b["cpu"] == "Core i7-12700 or Ryzen 7 7800X3D"
        assert b["ram"] == "16 GB RAM"
        assert b["gpu"] == ("GeForce RTX 2060 SUPER or Radeon RX 5700 XT "
                            "or Arc A770")
        assert b["directx"] == "Version 12"
        assert b["storage"] == "70 GB available space"
        assert b["os_legacy_flag"] is False

    def test_parsing_dos_requisitos_minimos(self):
        b = parse_requirements_block(CYBERPUNK_MINIMUM)
        assert b["cpu"] == "Core i7-6700 or Ryzen 5 1600"
        assert b["ram"] == "12 GB RAM"
        assert b["gpu"] == "GeForce GTX 1060 6GB"

    def test_item_sem_dois_pontos_vai_para_notas(self):
        b = parse_requirements_block(CYBERPUNK_RECOMMENDED)
        assert "64-bit processor" in (b["additional_notes"] or "")

    def test_raw_preservado_integralmente(self):
        b = parse_requirements_block(CYBERPUNK_RECOMMENDED)
        assert b["raw"] == CYBERPUNK_RECOMMENDED

    def test_formato_detectado_como_A(self):
        assert detect_markup_format(CYBERPUNK_RECOMMENDED) == "A"


class TestFormatoB:
    """Terraria: rotulo E valor dentro do mesmo <strong>.

    Um parser que assuma "valor = texto apos </strong>" devolve string vazia
    para TODOS estes campos. Este teste existe para impedir essa regressao.
    """

    def test_campos_nao_ficam_vazios(self):
        b = parse_requirements_block(TERRARIA_RECOMMENDED)
        for field in ("os", "cpu", "ram", "storage", "gpu"):
            assert b[field], f"campo {field} vazio no formato B"

    def test_valores_corretos(self):
        b = parse_requirements_block(TERRARIA_RECOMMENDED)
        assert b["os"] == "Windows 7, 8/8.1, 10"
        assert b["cpu"] == "Dual Core 3.0 Ghz"
        assert b["ram"] == "4GB"

    def test_sinonimo_hard_disk_space_mapeia_para_storage(self):
        b = parse_requirements_block(TERRARIA_RECOMMENDED)
        assert b["storage"] == "200MB"

    def test_sinonimo_video_card_mapeia_para_gpu(self):
        b = parse_requirements_block(TERRARIA_RECOMMENDED)
        assert b["gpu"] == "256mb Video Memory, capable of Shader Model 2.0+"

    def test_formato_detectado_como_B(self):
        assert detect_markup_format(TERRARIA_RECOMMENDED) == "B"

    def test_cabecalho_h2_nao_vira_campo(self):
        b = parse_requirements_block(TERRARIA_RECOMMENDED)
        rotulos = [u["label"].casefold() for u in b["unparsed_labels"]]
        assert "recommended" not in rotulos


class TestOsLegado:
    """O asterisco em `OS *` e sinal, nao ruido."""

    def test_asterisco_liga_a_flag(self):
        b = parse_requirements_block(PORTAL2_MINIMUM)
        assert b["os_legacy_flag"] is True
        assert b["os"] == "Windows 7 / Vista / XP"

    def test_sinonimo_hard_drive(self):
        b = parse_requirements_block(PORTAL2_MINIMUM)
        assert b["storage"] == "At least 7.6 GB of free space"

    def test_asterisco_tambem_no_recomendado(self):
        b = parse_requirements_block(HADES_RECOMMENDED)
        assert b["os_legacy_flag"] is True
        assert b["ram"] == "8 GB RAM"


class TestAusenciaDeRecomendados:
    """Item 8: ausencia de recomendados NAO pode derrubar o pipeline."""

    def test_recommended_ausente_gera_nulls_sem_excecao(self):
        out = parse_pc_requirements({"minimum": CYBERPUNK_MINIMUM})
        assert out["has_minimum_requirements"] is True
        assert out["has_recommended_requirements"] is False
        rec = out["pc_requirements"]["recommended"]
        assert rec["raw"] is None
        for field in ("os", "cpu", "ram", "gpu", "storage"):
            assert rec[field] is None

    def test_pc_requirements_totalmente_ausente(self):
        out = parse_pc_requirements(None)
        assert out["has_minimum_requirements"] is False
        assert out["has_recommended_requirements"] is False

    def test_pc_requirements_como_lista_vazia(self):
        # A Steam ja devolveu lista em vez de dict para alguns apps.
        out = parse_pc_requirements([])
        assert out["has_recommended_requirements"] is False

    def test_bloco_vazio_e_homogeneo(self):
        b = empty_block()
        assert set(b) == {"raw", "os", "cpu", "ram", "gpu", "directx", "storage",
                          "network", "sound_card", "additional_notes",
                          "os_legacy_flag", "unparsed_labels"}


class TestSeparacaoMinimoRecomendado:
    """Item 5: o minimo NUNCA pode vazar para o recomendado."""

    def test_valores_sao_independentes(self):
        out = parse_pc_requirements({"minimum": CYBERPUNK_MINIMUM,
                                     "recommended": CYBERPUNK_RECOMMENDED})
        mn = out["pc_requirements"]["minimum"]
        rc = out["pc_requirements"]["recommended"]
        assert mn["cpu"] == "Core i7-6700 or Ryzen 5 1600"
        assert rc["cpu"] == "Core i7-12700 or Ryzen 7 7800X3D"
        assert mn["cpu"] != rc["cpu"]
        assert mn["ram"] == "12 GB RAM" and rc["ram"] == "16 GB RAM"

    def test_minimo_nao_e_copiado_quando_recomendado_falta(self):
        out = parse_pc_requirements({"minimum": CYBERPUNK_MINIMUM})
        mn = out["pc_requirements"]["minimum"]
        rc = out["pc_requirements"]["recommended"]
        assert mn["cpu"] is not None
        assert rc["cpu"] is None
        assert rc["raw"] is None

    def test_blocos_nao_compartilham_objeto(self):
        out = parse_pc_requirements({"minimum": CYBERPUNK_MINIMUM,
                                     "recommended": CYBERPUNK_RECOMMENDED})
        mn = out["pc_requirements"]["minimum"]
        rc = out["pc_requirements"]["recommended"]
        assert mn is not rc
        assert mn["unparsed_labels"] is not rc["unparsed_labels"]


class TestParsingDeRamEArmazenamento:
    """Item 13 exige testes especificos de RAM e armazenamento."""

    @pytest.mark.parametrize("raw,esperado", [
        ("<ul><li><strong>Memory:</strong> 16 GB RAM</li></ul>", "16 GB RAM"),
        ("<ul><li><strong>Memory: 4GB</strong></li></ul>", "4GB"),
        ("<ul><li><strong>RAM:</strong> 8192 MB</li></ul>", "8192 MB"),
        ("<ul><li><strong>System Memory:</strong> 2 GB</li></ul>", "2 GB"),
    ])
    def test_ram(self, raw, esperado):
        assert parse_requirements_block(raw)["ram"] == esperado

    @pytest.mark.parametrize("raw,esperado", [
        ("<ul><li><strong>Storage:</strong> 70 GB available space</li></ul>",
         "70 GB available space"),
        ("<ul><li><strong>Hard Disk Space: 200MB</strong></li></ul>", "200MB"),
        ("<ul><li><strong>Hard Drive:</strong> 7.6 GB</li></ul>", "7.6 GB"),
        ("<ul><li><strong>Available Space:</strong> 50 GB</li></ul>", "50 GB"),
        ("<ul><li><strong>HDD:</strong> 1 GB</li></ul>", "1 GB"),
    ])
    def test_armazenamento(self, raw, esperado):
        assert parse_requirements_block(raw)["storage"] == esperado

    def test_nao_converte_para_numero(self):
        """Item 5: nesta etapa preserva texto, nao interpreta."""
        b = parse_requirements_block(
            "<ul><li><strong>Memory:</strong> 16 GB RAM</li></ul>")
        assert isinstance(b["ram"], str)


class TestRotulosDesconhecidos:
    """Rotulo nao mapeado nunca e descartado em silencio."""

    def test_sinonimo_supported_os(self):
        """Encontrado no piloto de 310 jogos: Left 4 Dead, Resident Evil 5."""
        raw = ('<ul class="bb_ul"><li><strong>Supported OS:</strong> '
               'Windows 7 32/64-bit / Vista 32/64 / XP</li></ul>')
        b = parse_requirements_block(raw)
        assert b["os"] == "Windows 7 32/64-bit / Vista 32/64 / XP"
        assert b["unparsed_labels"] == []

    def test_video_card_memory_nao_sobrescreve_gpu(self):
        """VRAM nao e o modelo da GPU; fica em unparsed_labels de proposito."""
        raw = ('<ul><li><strong>Graphics:</strong> GeForce GTX 650</li>'
               '<li><strong>Video Card Memory:</strong> 512 MB</li></ul>')
        b = parse_requirements_block(raw)
        assert b["gpu"] == "GeForce GTX 650"
        assert [u["label"] for u in b["unparsed_labels"]] == ["Video Card Memory"]

    def test_armazenamento_dentro_de_additional_notes(self):
        """Caso real (Call of Duty 1938090): a Steam nao publica o rotulo
        Storage; o dado esta em Additional Notes. storage=None e CORRETO, e a
        informacao fica preservada para o Data Wrangling."""
        raw = ('<ul class="bb_ul"><li><strong>Memory:</strong> 12 GB RAM</li>'
               '<li><strong>Additional Notes:</strong> SSD with 161 GB '
               'available space at launch</li></ul>')
        b = parse_requirements_block(raw)
        assert b["storage"] is None
        assert "161 GB" in b["additional_notes"]

    def test_rotulo_desconhecido_vai_para_unparsed_labels(self):
        raw = ('<ul class="bb_ul"><li><strong>Quantum Flux:</strong> 3 GW</li>'
               '<li><strong>Memory:</strong> 8 GB RAM</li></ul>')
        b = parse_requirements_block(raw)
        assert b["ram"] == "8 GB RAM"
        assert len(b["unparsed_labels"]) == 1
        assert b["unparsed_labels"][0]["label"] == "Quantum Flux"
        assert b["unparsed_labels"][0]["value"] == "3 GW"

    def test_rotulo_duplicado_nao_sobrescreve(self):
        raw = ('<ul><li><strong>Memory:</strong> 8 GB RAM</li>'
               '<li><strong>Memory:</strong> 16 GB RAM</li></ul>')
        b = parse_requirements_block(raw)
        assert b["ram"] == "8 GB RAM"
        assert any("duplicado" in u["label"] for u in b["unparsed_labels"])


class TestEntradasDegeneradas:
    def test_string_vazia(self):
        assert parse_requirements_block("")["raw"] is None

    def test_none(self):
        assert parse_requirements_block(None)["os"] is None

    def test_html_sem_li_usa_br(self):
        raw = ("<strong>Minimum:</strong><br>OS: Windows 10<br>"
               "Memory: 8 GB RAM<br>")
        b = parse_requirements_block(raw)
        assert b["os"] == "Windows 10"
        assert b["ram"] == "8 GB RAM"

    def test_entidades_html_decodificadas(self):
        raw = ("<ul><li><strong>OS:</strong> Windows&nbsp;10 "
               "Home &amp; Pro</li></ul>")
        assert parse_requirements_block(raw)["os"] == "Windows 10 Home & Pro"

    def test_rotulo_note_vai_para_notas(self):
        raw = ('<ul><li>Note: this game requires an internet connection for '
               'the following reasons and more: none</li></ul>')
        b = parse_requirements_block(raw)
        assert b["additional_notes"] is not None
        assert "internet connection" in b["additional_notes"]

    def test_prosa_longa_com_dois_pontos_nao_vira_rotulo(self):
        """Guarda contra tratar frase inteira como rotulo (limite de 45 chars)."""
        raw = ('<ul><li>This particular title happens to require the following '
               'additional consideration: a stable connection</li></ul>')
        b = parse_requirements_block(raw)
        assert b["unparsed_labels"] == []
        assert "stable connection" in (b["additional_notes"] or "")

    def test_formato_desconhecido(self):
        assert detect_markup_format("<div>sem estrutura</div>") == "UNKNOWN"
