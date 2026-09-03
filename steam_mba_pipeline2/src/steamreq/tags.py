"""Taxonomia oficial de tags e extracao das tags por jogo.

Duas responsabilidades:

1. Resolver nome -> tagid contra a taxonomia oficial (446 tags), com FALHA
   RUIDOSA se um nome do filters.yaml nao existir. Esse guard-rail existe
   porque, na FASE 1, supor que o tagid 4166 era "3D" (na verdade
   "Atmospheric") produziu um filtro silenciosamente errado.

2. Extrair as 20 tags do HTML da pagina do jogo. O appdetails NAO retorna
   tags: verificado em 9 jogos. `genres` e `categories` nao sao tags.
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass

from .logging_setup import get_logger, utc_now_iso
from .storage import read_json, write_json_atomic

log = get_logger("tags")

# InitAppTagModal( 1091500, [{"tagid":4115,"name":"Cyberpunk"}, ...], ...
_TAG_MODAL_RE = re.compile(
    r"InitAppTagModal\(\s*\d+\s*,\s*(\[.*?\])\s*,", re.DOTALL)


class TagResolutionError(RuntimeError):
    """Uma tag do filters.yaml nao existe na taxonomia oficial."""


@dataclass
class Taxonomy:
    """Snapshot versionado da taxonomia oficial de tags."""

    by_name: dict[str, int]
    by_id: dict[int, str]
    fetched_at: str

    @property
    def size(self) -> int:
        return len(self.by_name)

    # -- construcao --------------------------------------------------------
    @classmethod
    def from_payload(cls, payload: dict, fetched_at: str | None = None) -> Taxonomy:
        tags = (payload.get("response") or {}).get("tags") or payload.get("tags")
        if not isinstance(tags, list) or not tags:
            raise TagResolutionError(
                "resposta da taxonomia sem lista de tags utilizavel")
        by_name: dict[str, int] = {}
        by_id: dict[int, str] = {}
        for item in tags:
            name, tid = item.get("name"), item.get("tagid")
            if name is None or tid is None:
                continue
            by_name[str(name)] = int(tid)
            by_id[int(tid)] = str(name)
        if not by_name:
            raise TagResolutionError("nenhuma tag valida na resposta da taxonomia")
        return cls(by_name=by_name, by_id=by_id,
                   fetched_at=fetched_at or utc_now_iso())

    @classmethod
    def fetch(cls, client, settings, *, cache_dir: str | None = None,  # noqa: ANN001
              use_cache: bool = True) -> Taxonomy:
        """Baixa a taxonomia; usa cache local se disponivel e permitido."""
        cache_path = os.path.join(cache_dir, "tag_taxonomy.json") if cache_dir else None
        if use_cache and cache_path:
            cached = read_json(cache_path)
            if cached:
                tax = cls.from_payload(cached.get("payload", cached),
                                       cached.get("fetched_at"))
                log.info("[TAGS]      taxonomia carregada do cache | %d tags | "
                         "coletada em %s", tax.size, tax.fetched_at)
                return tax

        payload = client.get_json("api", settings.endpoints["tag_list"],
                                  {"language": "english"})
        tax = cls.from_payload(payload)
        if cache_path:
            write_json_atomic(cache_path,
                              {"fetched_at": tax.fetched_at, "payload": payload})
        log.info("[TAGS]      %d tags oficiais baixadas da Steam", tax.size)
        return tax

    # -- resolucao ---------------------------------------------------------
    def resolve(self, name: str) -> int | None:
        if name in self.by_name:
            return self.by_name[name]
        # Tolerancia apenas a caixa/espacos; nunca a "quase igual".
        norm = name.strip().casefold()
        for known, tid in self.by_name.items():
            if known.strip().casefold() == norm:
                return tid
        return None

    def resolve_all(self, names: list[str], *, strict: bool = True
                    ) -> tuple[dict[str, int], list[str]]:
        """Resolve uma lista de nomes. Em modo strict, aborta se faltar alguma.

        Excecao deliberada: "Homemade" foi confirmada como inexistente na
        taxonomia oficial (item 2.13), e por isso nao consta do filters.yaml.
        Qualquer outra ausencia indica erro de configuracao e deve abortar.
        """
        resolved: dict[str, int] = {}
        missing: list[str] = []
        for name in names:
            tid = self.resolve(name)
            if tid is None:
                missing.append(name)
            else:
                resolved[name] = tid
        if missing and strict:
            raise TagResolutionError(
                "tags inexistentes na taxonomia oficial da Steam: "
                f"{missing}. Corrija config/filters.yaml. Prosseguir produziria "
                "um filtro silenciosamente vazio."
            )
        return resolved, missing


# --- extracao das tags por jogo --------------------------------------------

def extract_tags_from_html(html: str) -> list[dict]:
    """Extrai as ~20 tags de InitAppTagModal.

    Retorna [] se o bloco nao existir — sinal de possivel mudanca de estrutura
    da pagina (item 10), tratado como PARTIAL_NO_TAGS pelo scraper.
    """
    match = _TAG_MODAL_RE.search(html)
    if not match:
        return []
    try:
        tags = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        log.warning("InitAppTagModal encontrado mas com JSON invalido: %s", exc)
        return []
    out = []
    for item in tags:
        if isinstance(item, dict) and "tagid" in item and "name" in item:
            out.append({"tagid": int(item["tagid"]), "name": str(item["name"])})
    return out


def extract_tagids_from_search_row(row_html: str) -> list[int]:
    """Le data-ds-tagids de uma linha do /search/results/.

    Apenas 7-8 tagids (subconjunto das 20). Usado como sinal auxiliar no
    discovery, NUNCA como fonte primaria de tags — produziria falsos negativos,
    proibidos pelo item 2.16.
    """
    match = re.search(r'data-ds-tagids="(\[[^\]]*\])"', row_html)
    if not match:
        return []
    try:
        return [int(t) for t in json.loads(match.group(1))]
    except (json.JSONDecodeError, ValueError, TypeError):
        return []
