"""Carga e validacao das configuracoes YAML."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any

import yaml

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))


class ConfigError(RuntimeError):
    """Configuracao ausente, malformada ou incoerente."""


def _read_yaml(path: str) -> dict[str, Any]:
    if not os.path.exists(path):
        raise ConfigError(f"arquivo de configuracao nao encontrado: {path}")
    with open(path, encoding="utf-8") as fh:
        data = yaml.safe_load(fh)
    if not isinstance(data, dict):
        raise ConfigError(f"conteudo invalido em {path}: esperado mapeamento YAML")
    return data


@dataclass
class Settings:
    """Configuracao operacional (settings.yaml)."""

    raw: dict[str, Any] = field(repr=False)
    scraper_version: str
    http: dict[str, Any]
    rate_limits: dict[str, Any]
    endpoints: dict[str, str]
    discovery: dict[str, Any]
    scrape: dict[str, Any]
    paths: dict[str, str]
    logging_cfg: dict[str, Any]

    @classmethod
    def load(cls, path: str | None = None) -> Settings:
        path = path or os.path.join(PROJECT_ROOT, "config", "settings.yaml")
        d = _read_yaml(path)
        required = ["scraper_version", "http", "rate_limits", "endpoints",
                    "discovery", "scrape", "paths"]
        missing = [k for k in required if k not in d]
        if missing:
            raise ConfigError(f"chaves ausentes em settings.yaml: {missing}")
        return cls(
            raw=d,
            scraper_version=str(d["scraper_version"]),
            http=d["http"],
            rate_limits=d["rate_limits"],
            endpoints=d["endpoints"],
            discovery=d["discovery"],
            scrape=d["scrape"],
            paths=d["paths"],
            logging_cfg=d.get("logging", {}),
        )

    def path(self, key: str) -> str:
        """Resolve um caminho de `paths` para absoluto, criando o diretorio."""
        if key not in self.paths:
            raise ConfigError(f"caminho '{key}' nao definido em settings.yaml")
        p = os.path.join(PROJECT_ROOT, self.paths[key])
        os.makedirs(p, exist_ok=True)
        return p


@dataclass
class Filters:
    """Metodologia de filtragem (filters.yaml, item 2 do projeto)."""

    raw: dict[str, Any] = field(repr=False)
    filter_version: str
    app_type_allowed: list[str]
    require_windows: bool
    min_date: str
    exclude_coming_soon: bool
    exclude_early_access: bool
    positive_strong: list[str]
    positive_secondary: list[str]
    negative_2d: list[str]
    neutral_never_exclude: list[str]
    exclude_at_source: list[str]
    negative_candidates: list[str]
    never_exclude: list[str]
    review_metric: str
    review_min_threshold: int | None
    review_candidate_thresholds: list[int]

    @classmethod
    def load(cls, path: str | None = None) -> Filters:
        path = path or os.path.join(PROJECT_ROOT, "config", "filters.yaml")
        d = _read_yaml(path)
        if "filter_version" not in d:
            raise ConfigError("filters.yaml sem filter_version: a rastreabilidade "
                              "do item 2.17 depende dele")
        tags = d.get("tags", {})
        rel = d.get("release", {})
        rev = d.get("reviews", {})

        obj = cls(
            raw=d,
            filter_version=str(d["filter_version"]),
            app_type_allowed=list(d.get("app_type", {}).get("allowed", ["game"])),
            require_windows=bool(d.get("platform", {}).get("require_windows", True)),
            min_date=str(rel.get("min_date", "2005-01-01")),
            exclude_coming_soon=bool(rel.get("exclude_coming_soon", True)),
            exclude_early_access=bool(rel.get("exclude_early_access", False)),
            positive_strong=list(tags.get("positive_strong", [])),
            positive_secondary=list(tags.get("positive_secondary", [])),
            negative_2d=list(tags.get("negative_2d", [])),
            neutral_never_exclude=list(tags.get("neutral_never_exclude", [])),
            exclude_at_source=list(tags.get("exclude_at_source", [])),
            negative_candidates=list(tags.get("negative_candidates", [])),
            never_exclude=list(tags.get("never_exclude", [])),
            review_metric=str(rev.get("metric", "recommendations_total")),
            review_min_threshold=rev.get("min_threshold"),
            review_candidate_thresholds=list(rev.get("candidate_thresholds",
                                                     [0, 100, 500, 1000, 5000])),
        )
        obj._validate_methodology()
        return obj

    def _validate_methodology(self) -> None:
        """Impede configuracoes que violem a metodologia do projeto."""
        # Item 2.12: estas tags nao podem ser usadas como exclusao.
        forbidden = {t.lower() for t in self.never_exclude}
        for bucket_name, bucket in (("exclude_at_source", self.exclude_at_source),
                                    ("negative_2d", self.negative_2d)):
            for tag in bucket:
                if tag.lower() in forbidden:
                    raise ConfigError(
                        f"violacao do item 2.12: a tag '{tag}' esta em "
                        f"'{bucket_name}' mas tambem em 'never_exclude'"
                    )
        # Item 2.9: 2.5D nunca pode ser criterio de exclusao.
        neutral = {t.lower() for t in self.neutral_never_exclude}
        for tag in self.exclude_at_source + self.negative_2d:
            if tag.lower() in neutral:
                raise ConfigError(
                    f"violacao do item 2.9: a tag neutra '{tag}' esta sendo "
                    f"usada como exclusao"
                )
        if not self.positive_strong:
            raise ConfigError("positive_strong vazio: o discovery nao teria "
                              "nenhuma tag para consultar")

    @property
    def all_referenced_tags(self) -> list[str]:
        """Todas as tags citadas no filtro, para resolucao contra a taxonomia."""
        seen: dict[str, None] = {}
        for bucket in (self.positive_strong, self.positive_secondary,
                       self.negative_2d, self.neutral_never_exclude,
                       self.exclude_at_source, self.negative_candidates,
                       self.never_exclude):
            for tag in bucket:
                seen.setdefault(tag, None)
        return list(seen)


def load_all(settings_path: str | None = None,
             filters_path: str | None = None) -> tuple[Settings, Filters]:
    return Settings.load(settings_path), Filters.load(filters_path)
