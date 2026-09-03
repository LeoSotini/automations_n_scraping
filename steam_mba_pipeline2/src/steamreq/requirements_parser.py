"""Parsing de pc_requirements (item 5 do projeto — PRIORIDADE MAXIMA).

Contexto empirico da FASE 1 (docs/FASE1_investigacao.md, secao 3):

* `pc_requirements` e um dict de STRINGS HTML: {"minimum": "...", "recommended": "..."}.
  `recommended` PODE ESTAR AUSENTE (3 de 9 jogos amostrados).
* Existem ao menos DOIS formatos de markup incompativeis:
    Formato A (Cyberpunk):  <li><strong>Processor:</strong> Core i7-12700
    Formato B (Terraria):   <li><strong>Processor: Dual Core 3.0 Ghz</strong>
  Um parser que assuma "valor = texto apos </strong>" retorna string VAZIA
  para todos os campos do Formato B.
* Os rotulos variam para o mesmo conceito: Storage/Hard Disk Space,
  Graphics/Video Card, e `OS *` (asterisco = SO legado sem suporte).

ESTRATEGIA: em vez de depender da posicao das tags <strong>, extraimos o TEXTO
de cada <li> e dividimos no primeiro ':'. Isso trata A e B de forma uniforme,
sem depender do markup — que e justamente a parte instavel.

ESCOPO DELIBERADAMENTE LIMITADO (item 5): extrai texto. NAO converte "16 GB RAM"
para 16, NAO mapeia GPU para benchmark. Normalizacao de hardware pertence ao
Data Wrangling. Aqui a prioridade e fidelidade, nao interpretacao.
"""
from __future__ import annotations

import html as html_module
import re
from typing import Any

from .logging_setup import get_logger

log = get_logger("reqparser")

# Campos estruturados do schema. Ordem estavel para o export.
FIELDS = ("os", "cpu", "ram", "gpu", "directx", "storage", "network",
          "sound_card", "additional_notes")

# Mapa de sinonimos de rotulos -> campo canonico.
# Chaves normalizadas: minusculas, sem pontuacao/asterisco/espacos extras.
LABEL_SYNONYMS: dict[str, str] = {
    # SO
    "os": "os", "operating system": "os", "os version": "os", "system": "os",
    "supported os": "os", "supported operating system": "os",
    # CPU
    "processor": "cpu", "cpu": "cpu", "processador": "cpu",
    # RAM
    "memory": "ram", "ram": "ram", "memoria": "ram", "system memory": "ram",
    # GPU
    "graphics": "gpu", "video card": "gpu", "graphics card": "gpu",
    "video": "gpu", "gpu": "gpu", "graphics hardware": "gpu",
    "video memory": "gpu",
    # NAO mapeados de proposito, para nao sobrescrever campos canonicos nem
    # inventar semantica (auditado em escala no piloto de 310 jogos):
    #   "Video Card Memory" -> e VRAM, nao o modelo da GPU
    #   "VR Support", "Display", "Peripherals", "Input" -> fora do escopo do
    #   item 4/5 do projeto
    # Todos permanecem visiveis em `unparsed_labels`.
    # DirectX
    "directx": "directx", "directx version": "directx", "direct x": "directx",
    # Armazenamento
    "storage": "storage", "hard disk space": "storage", "hard drive": "storage",
    "hdd": "storage", "available space": "storage", "disk space": "storage",
    "hard disk": "storage", "hard drive space": "storage",
    "free disk space": "storage",
    # Rede
    "network": "network", "internet connection": "network",
    "internet": "network", "network connection": "network",
    # Som
    "sound card": "sound_card", "sound": "sound_card", "audio": "sound_card",
    # Notas
    "additional notes": "additional_notes", "notes": "additional_notes",
    "note": "additional_notes", "other": "additional_notes",
    "other requirements": "additional_notes",
    "additional": "additional_notes",
}

_LI_RE = re.compile(r"<li[^>]*>(.*?)(?:</li>|(?=<li)|$)", re.DOTALL | re.IGNORECASE)
_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")
# Formato A: </strong> seguido de conteudo textual nao vazio.
_FORMAT_A_RE = re.compile(r"</strong>\s*(?!<)[^<\s][^<]*", re.IGNORECASE)
# Formato B: rotulo E valor dentro do mesmo <strong>.
_FORMAT_B_RE = re.compile(r"<strong[^>]*>\s*[^<:]{1,40}:\s*\S[^<]*</strong>",
                          re.IGNORECASE)


def _text_of(fragment: str) -> str:
    """Remove tags, converte entidades e colapsa espacos."""
    # <br> vira espaco para nao colar palavras de linhas diferentes.
    fragment = re.sub(r"<br\s*/?>", " ", fragment, flags=re.IGNORECASE)
    text = _TAG_RE.sub("", fragment)
    text = html_module.unescape(text)
    text = text.replace("\xa0", " ")
    return _WS_RE.sub(" ", text).strip()


def _normalize_label(label: str) -> tuple[str, bool]:
    """Normaliza um rotulo. Retorna (chave_normalizada, tinha_asterisco).

    O asterisco em `OS *` marca SO legado sem suporte oficial. E SINAL, nao
    ruido: viaja para os_legacy_flag em vez de ser descartado.
    """
    had_asterisk = "*" in label
    cleaned = label.replace("*", " ")
    cleaned = re.sub(r"[^\w\s&+/-]", " ", cleaned, flags=re.UNICODE)
    cleaned = _WS_RE.sub(" ", cleaned).strip().casefold()
    return cleaned, had_asterisk


def detect_markup_format(raw_html: str) -> str:
    """Identifica o formato do markup: 'A', 'B' ou 'UNKNOWN'."""
    if not raw_html:
        return "UNKNOWN"
    has_a = bool(_FORMAT_A_RE.search(raw_html))
    has_b = bool(_FORMAT_B_RE.search(raw_html))
    if has_a and has_b:
        return "A"      # markup misto: o caminho A cobre os itens canonicos
    if has_a:
        return "A"
    if has_b:
        return "B"
    return "UNKNOWN"


def _split_items(raw_html: str) -> list[str]:
    """Divide o bloco em itens. Usa <li>; se nao houver, cai para <br>."""
    items = [m.group(1) for m in _LI_RE.finditer(raw_html)]
    if items:
        return items
    body = re.split(r"</?ul[^>]*>", raw_html, flags=re.IGNORECASE)
    body = body[1] if len(body) > 1 else raw_html
    return [p for p in re.split(r"<br\s*/?>", body, flags=re.IGNORECASE)
            if _text_of(p)]


def _strip_section_header(raw_html: str) -> str:
    """Remove o cabecalho 'Minimum:'/'RECOMMENDED' antes dos itens.

    Cobre as duas formas observadas: <strong>Recommended:</strong> e
    <h2 class="bb_tag"><strong>RECOMMENDED</strong></h2>.
    """
    return re.sub(
        r"^\s*(?:<h2[^>]*>)?\s*<strong>\s*(?:minimum|recommended|minimo|"
        r"recomendado)\s*:?\s*</strong>\s*(?:</h2>)?\s*(?:<br\s*/?>)?",
        "", raw_html, flags=re.IGNORECASE)


def empty_block() -> dict[str, Any]:
    """Bloco vazio homogeneo. Item 5: nao forcar valores, usar null."""
    block: dict[str, Any] = {"raw": None}
    block.update({f: None for f in FIELDS})
    block["os_legacy_flag"] = False
    block["unparsed_labels"] = []
    return block


def parse_requirements_block(raw_html: str | None) -> dict[str, Any]:
    """Converte um bloco HTML de requisitos em campos estruturados.

    O `raw` original e SEMPRE preservado integralmente (item 5), permitindo
    corrigir erros de parsing depois sem re-raspar a Steam.
    """
    block = empty_block()
    if not raw_html or not str(raw_html).strip():
        return block

    raw_html = str(raw_html)
    block["raw"] = raw_html                      # fidelidade acima de tudo

    body = _strip_section_header(raw_html)
    notes: list[str] = []

    for item in _split_items(body):
        text = _text_of(item)
        if not text:
            continue

        if ":" not in text:
            # Ex.: "Requires a 64-bit processor and operating system".
            notes.append(text)
            continue

        label_part, _, value_part = text.partition(":")
        value = value_part.strip(" .;")
        key, had_asterisk = _normalize_label(label_part)

        # Rotulo longo demais para ser rotulo: e prosa com dois-pontos.
        if len(key) > 45:
            notes.append(text)
            continue

        field = LABEL_SYNONYMS.get(key)
        if field is None:
            block["unparsed_labels"].append(
                {"label": label_part.strip(), "value": value or None})
            continue

        if field == "os" and had_asterisk:
            block["os_legacy_flag"] = True

        if not value:
            continue
        if block[field] is None:
            block[field] = value
        elif field == "additional_notes":
            block[field] = f"{block[field]} | {value}"
        else:
            # Rotulo repetido (ocorre em paginas mal formatadas): preserva o
            # primeiro e registra o extra para auditoria, sem sobrescrever.
            block["unparsed_labels"].append(
                {"label": f"{label_part.strip()} (duplicado)", "value": value})

    if notes:
        merged = " | ".join(notes)
        block["additional_notes"] = (
            f"{block['additional_notes']} | {merged}"
            if block["additional_notes"] else merged)

    return block


def parse_pc_requirements(pc_requirements: Any) -> dict[str, Any]:
    """Processa o `pc_requirements` do appdetails.

    Garantias, exigidas pelo item 5:
    * `minimum` e `recommended` sao parseados de forma INDEPENDENTE;
    * nenhum valor de `minimum` e copiado para `recommended` em qualquer
      circunstancia — sao dicionarios distintos, produzidos por chamadas
      distintas, a partir de strings distintas;
    * ausencia de `recommended` produz bloco com null, nao excecao (item 8).
    """
    minimum_raw: str | None = None
    recommended_raw: str | None = None

    # A Steam ja devolveu `pc_requirements` como lista vazia para alguns apps.
    if isinstance(pc_requirements, dict):
        minimum_raw = pc_requirements.get("minimum") or None
        recommended_raw = pc_requirements.get("recommended") or None

    minimum = parse_requirements_block(minimum_raw)
    recommended = parse_requirements_block(recommended_raw)

    fmt_source = recommended_raw or minimum_raw
    return {
        "pc_requirements": {"minimum": minimum, "recommended": recommended},
        "has_minimum_requirements": minimum_raw is not None,
        "has_recommended_requirements": recommended_raw is not None,
        "requirements_markup_format": detect_markup_format(fmt_source or ""),
    }
