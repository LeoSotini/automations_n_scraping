"""Persistencia: escrita atomica, checkpoints e ledgers (itens 7, 10 e 11).

Garantia central: nenhuma interrupcao (Ctrl+C, queda de energia, erro de disco)
pode deixar um checkpoint corrompido. Toda escrita vai para <arquivo>.tmp e
depois usa os.replace(), que e atomico no mesmo volume em Windows e POSIX.
"""
from __future__ import annotations

import json
import os
import tempfile
import time
from enum import StrEnum
from typing import Any, Iterable

from .logging_setup import get_logger, utc_now_iso

log = get_logger("storage")


class Status(StrEnum):
    """Estados possiveis de um app no checkpoint (item 7)."""

    PENDING = "PENDING"
    IN_PROGRESS = "IN_PROGRESS"
    COMPLETE = "COMPLETE"
    PARTIAL_NO_TAGS = "PARTIAL_NO_TAGS"
    FAILED = "FAILED"
    EXCLUDED_BY_FILTER = "EXCLUDED_BY_FILTER"

    @property
    def is_terminal_success(self) -> bool:
        return self in (Status.COMPLETE, Status.PARTIAL_NO_TAGS)


# --- IO atomico -------------------------------------------------------------

_REPLACE_RETRIES = 5
_REPLACE_BACKOFF_S = 0.05


def _replace_with_retry(tmp: str, path: str) -> None:
    """os.replace resiliente a bloqueios transitorios do Windows.

    Em Windows, os.replace pode falhar com PermissionError (WinError 5/32)
    quando outro processo mantem o destino aberto por um instante — indexador
    do sistema, antivirus, backup de nuvem (o projeto vive numa pasta OneDrive).
    Observado empiricamente durante os testes, com centenas de gravacoes
    consecutivas de checkpoint.

    A operacao permanece atomica: ou a troca ocorre, ou o arquivo anterior
    continua intacto. Apenas insistimos por alguns instantes antes de desistir.
    """
    for attempt in range(1, _REPLACE_RETRIES + 1):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            if attempt == _REPLACE_RETRIES:
                raise
            time.sleep(_REPLACE_BACKOFF_S * attempt)
            log.debug("os.replace bloqueado em %s; tentativa %d/%d",
                      path, attempt + 1, _REPLACE_RETRIES)


def write_json_atomic(path: str, data: Any, *, indent: int | None = None) -> None:
    """Grava JSON UTF-8 de forma atomica (item 8: encoding e serializacao)."""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    directory = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=directory, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=indent)
            fh.flush()
            os.fsync(fh.fileno())
        _replace_with_retry(tmp, path)
    except BaseException:
        # Nunca deixa .tmp orfao nem sobrescreve o arquivo bom.
        with_suppress_unlink(tmp)
        raise


def with_suppress_unlink(path: str) -> None:
    try:
        if os.path.exists(path):
            os.unlink(path)
    except OSError:  # pragma: no cover - best effort
        pass


def read_json(path: str, default: Any = None) -> Any:
    """Le JSON tolerando ausencia e corrupcao (item 10: JSON invalido)."""
    if not os.path.exists(path):
        return default
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        log.error("arquivo JSON corrompido em %s (%s); tratando como ausente",
                  path, exc)
        return default


# --- cache de payloads brutos (data/raw, imutavel) -------------------------

class RawStore:
    """Armazena os payloads brutos. Habilita reprocessamento offline.

    Consequencia de projeto: alterar filtros ou corrigir o parser NAO exige
    re-raspar a Steam (itens 5 e 12).
    """

    def __init__(self, appdetails_dir: str, tags_dir: str) -> None:
        self.appdetails_dir = appdetails_dir
        self.tags_dir = tags_dir
        os.makedirs(appdetails_dir, exist_ok=True)
        os.makedirs(tags_dir, exist_ok=True)

    def appdetails_path(self, app_id: int) -> str:
        return os.path.join(self.appdetails_dir, f"{app_id}.json")

    def tags_path(self, app_id: int) -> str:
        return os.path.join(self.tags_dir, f"{app_id}.tags.json")

    def has_appdetails(self, app_id: int) -> bool:
        return os.path.exists(self.appdetails_path(app_id))

    def has_tags(self, app_id: int) -> bool:
        return os.path.exists(self.tags_path(app_id))

    def save_appdetails(self, app_id: int, payload: dict) -> None:
        write_json_atomic(self.appdetails_path(app_id), payload)

    def load_appdetails(self, app_id: int) -> dict | None:
        return read_json(self.appdetails_path(app_id))

    def save_tags(self, app_id: int, payload: dict) -> None:
        write_json_atomic(self.tags_path(app_id), payload)

    def load_tags(self, app_id: int) -> dict | None:
        return read_json(self.tags_path(app_id))

    def known_app_ids(self) -> list[int]:
        out = []
        for fn in os.listdir(self.appdetails_dir):
            if fn.endswith(".json"):
                try:
                    out.append(int(fn[:-5]))
                except ValueError:
                    continue
        return sorted(out)


# --- checkpoint -------------------------------------------------------------

class Checkpoint:
    """Estado por app_id, retomavel e idempotente (itens 7 e 11).

    Sabe, conforme exigido: pendentes, concluidos, excluidos pelo filtro, que
    falharam, numero de tentativas, tipo do ultimo erro e timestamp da ultima
    tentativa.
    """

    def __init__(self, path: str, *, filter_version: str, scraper_version: str,
                 flush_every: int = 25) -> None:
        self.path = path
        self.flush_every = max(1, flush_every)
        self._dirty = 0
        state = read_json(path, default=None) or {}
        self.meta: dict[str, Any] = state.get("meta") or {
            "started_at": utc_now_iso(),
            "filter_version": filter_version,
            "scraper_version": scraper_version,
        }
        self.apps: dict[str, dict[str, Any]] = state.get("apps") or {}

        # Item 12: mudanca de versao de filtro nao pode passar em silencio.
        prev_fv = self.meta.get("filter_version")
        if prev_fv and prev_fv != filter_version:
            log.warning(
                "checkpoint foi criado com filter_version=%s mas a configuracao "
                "atual e %s. Os registros ja coletados permanecem validos "
                "(o raw e imutavel), mas o estagio de filtro deve ser reexecutado.",
                prev_fv, filter_version)
        self.meta["filter_version"] = filter_version
        self.meta["scraper_version"] = scraper_version

    # -- consultas ---------------------------------------------------------
    def get(self, app_id: int) -> dict[str, Any] | None:
        return self.apps.get(str(app_id))

    def status_of(self, app_id: int) -> Status:
        entry = self.get(app_id)
        if not entry:
            return Status.PENDING
        try:
            return Status(entry.get("status", "PENDING"))
        except ValueError:
            return Status.PENDING

    def attempts_of(self, app_id: int) -> int:
        entry = self.get(app_id)
        return int(entry.get("attempts", 0)) if entry else 0

    def is_done(self, app_id: int) -> bool:
        """Idempotencia: apps concluidos sao pulados sem requisicao de rede."""
        return self.status_of(app_id).is_terminal_success

    def counts(self) -> dict[str, int]:
        out = {s.value: 0 for s in Status}
        for entry in self.apps.values():
            out[entry.get("status", "PENDING")] = out.get(
                entry.get("status", "PENDING"), 0) + 1
        return out

    def app_ids_with_status(self, *statuses: Status) -> list[int]:
        wanted = {s.value for s in statuses}
        return sorted(int(k) for k, v in self.apps.items()
                      if v.get("status") in wanted)

    # -- mutacoes ----------------------------------------------------------
    def mark(self, app_id: int, status: Status, *, error: str | None = None,
             increment_attempt: bool = False, **extra: Any) -> None:
        key = str(app_id)
        entry = self.apps.setdefault(key, {"attempts": 0})
        if increment_attempt:
            entry["attempts"] = int(entry.get("attempts", 0)) + 1
        entry["status"] = status.value
        entry["last_error"] = error
        entry["last_attempt_at"] = utc_now_iso()
        entry.update(extra)
        self._dirty += 1
        if self._dirty >= self.flush_every:
            self.flush()

    def register_pending(self, app_ids: Iterable[int]) -> int:
        """Registra candidatos como PENDING sem sobrescrever estados existentes."""
        added = 0
        for app_id in app_ids:
            if str(app_id) not in self.apps:
                self.apps[str(app_id)] = {"status": Status.PENDING.value,
                                          "attempts": 0, "last_error": None,
                                          "last_attempt_at": None}
                added += 1
        if added:
            self._dirty += added
            self.flush()
        return added

    def reset_failed(self) -> int:
        """Prepara reprocessamento de falhas (item 11: retry-failed)."""
        n = 0
        for entry in self.apps.values():
            if entry.get("status") == Status.FAILED.value:
                entry["status"] = Status.PENDING.value
                n += 1
        if n:
            self.flush()
        return n

    def flush(self) -> None:
        self.meta["last_updated"] = utc_now_iso()
        write_json_atomic(self.path, {"meta": self.meta, "apps": self.apps})
        self._dirty = 0
        log.debug("checkpoint gravado em %s (%d apps)", self.path, len(self.apps))


# --- ledgers de exclusao (item 2.17) ---------------------------------------

def append_ledger(path: str, records: list[dict[str, Any]], *,
                  filter_version: str) -> int:
    """Acrescenta registros a um ledger de exclusao, deduplicando por app_id.

    O item 2.17 proibe descartar registros sem registrar o motivo. Os ledgers
    sao o instrumento que torna as exclusoes auditaveis e reversiveis.
    """
    existing = read_json(path, default=None) or {"meta": {}, "records": []}
    by_id = {int(r["app_id"]): r for r in existing.get("records", [])}
    added = 0
    for rec in records:
        aid = int(rec["app_id"])
        if aid not in by_id:
            by_id[aid] = rec
            added += 1
    payload = {
        "meta": {"filter_version": filter_version,
                 "last_updated": utc_now_iso(),
                 "total": len(by_id)},
        "records": [by_id[k] for k in sorted(by_id)],
    }
    write_json_atomic(path, payload, indent=2)
    return added
