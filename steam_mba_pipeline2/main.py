"""CLI do pipeline de coleta de requisitos da Steam (item 11 do projeto).

    python main.py tags                # taxonomia oficial + resolve filters.yaml
    python main.py discover            # estagio 1 -> candidates.json
    python main.py scrape [--limit N]  # estagio 2, retomavel
    python main.py resume              # retoma o scrape
    python main.py retry-failed        # reprocessa apenas as falhas
    python main.py filter              # estagio 3, offline sobre data/raw
    python main.py validate            # estagio 5
    python main.py export [--flat]     # estagio 6 -> dataset.json
    python main.py analyze-thresholds  # item 2.14
    python main.py audit-bias          # decisao D1 (exige STEAM_API_KEY)
    python main.py status              # resumo do checkpoint
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from steamreq import __version__  # noqa: E402
from steamreq.config import ConfigError, load_all  # noqa: E402
from steamreq.discovery import Discovery, load_candidate_ids  # noqa: E402
from steamreq.export import (analyze_thresholds, export_dataset,  # noqa: E402
                            run_filter_stage)
from steamreq.logging_setup import get_logger, log_path, setup  # noqa: E402
from steamreq.steam_client import SteamClient  # noqa: E402
from steamreq.storage import Checkpoint, RawStore, Status  # noqa: E402
from steamreq.tags import TagResolutionError, Taxonomy  # noqa: E402
from steamreq.validators import validate_dataset  # noqa: E402

log = get_logger("cli")

# Amostra heterogenea da FASE 3, escolhida para cobrir os casos de borda
# efetivamente observados na FASE 1 (nao sao exemplos hipoteticos).
PROTOTYPE_SAMPLE: list[tuple[int, str]] = [
    (220,     "Half-Life 2 (2004: deve ser excluido por BEFORE_2005)"),
    (620,     "Portal 2 (2011: SEM requisitos recomendados)"),
    (730,     "Counter-Strike 2 (F2P, SEM recomendados, item 2.12)"),
    (105600,  "Terraria (FORMATO B de markup + tags 2D)"),
    (271590,  "GTA V Legacy (age gate que NAO bloqueia)"),
    (1091500, "Cyberpunk 2077 (FORMATO A canonico, AAA recente)"),
    (1245620, "ELDEN RING (AAA 2022, Souls-like)"),
    (1174180, "Red Dead Redemption 2 (AAA 2019)"),
    (1145360, "Hades (Indie + 2.5D: teste do item 2.9/2.10)"),
    (2138330, "Cyberpunk: Phantom Liberty (DLC: deve dar NOT_GAME)"),
    (339800,  "HuniePop (age gate COM login: PARTIAL_NO_TAGS)"),
    (999999999, "app_id inexistente (HTTP 200 com success:false)"),
    (292030,  "The Witcher 3 (2015, AAA)"),
    (582010,  "Monster Hunter World (2018)"),
    (367520,  "Hollow Knight (Indie 2D: teste de exclusao)"),
    (413150,  "Stardew Valley (Pixel Graphics + Indie)"),
    (48700,   "Mount & Blade: Warband (2010, antigo)"),
]


def _bootstrap(args: argparse.Namespace):
    settings, filters = load_all(args.settings, args.filters)
    setup(settings.path("logs"),
          console_level="DEBUG" if args.verbose else
          settings.logging_cfg.get("console_level", "INFO"),
          file_level=settings.logging_cfg.get("file_level", "DEBUG"))
    log.info("=" * 78)
    log.info("steam-req-pipeline v%s | scraper=%s | filter_version=%s",
             __version__, settings.scraper_version, filters.filter_version)
    log.info("log persistente: %s", log_path())
    log.info("=" * 78)
    return settings, filters


def _load_taxonomy(client, settings, *, use_cache: bool = True) -> Taxonomy:
    return Taxonomy.fetch(client, settings,
                          cache_dir=settings.path("raw_taxonomy"),
                          use_cache=use_cache)


def _resolve_and_report(taxonomy: Taxonomy, filters) -> None:
    """Resolve as tags do filtro. Aborta se alguma nao existir."""
    names = filters.all_referenced_tags
    resolved, missing = taxonomy.resolve_all(names, strict=False)
    log.info("[TAGS]      %d tags oficiais | %d/%d tags do filtro resolvidas",
             taxonomy.size, len(resolved), len(names))
    if missing:
        # O item 2.13 confirmou que "Homemade" nao existe; qualquer ausencia
        # aqui e erro de configuracao e deve abortar (nunca filtro vazio).
        raise TagResolutionError(
            f"tags inexistentes na taxonomia oficial: {missing}. "
            "Corrija config/filters.yaml.")
    for name in filters.positive_strong:
        log.debug("[TAGS]      positiva forte %-22s -> %d", name, resolved[name])


# --- comandos ---------------------------------------------------------------

def cmd_tags(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    with SteamClient(settings) as client:
        taxonomy = _load_taxonomy(client, settings, use_cache=not args.refresh)
        _resolve_and_report(taxonomy, filters)
    log.info("[TAGS]      OK: a metodologia de tags esta integralmente "
             "resolvida contra a taxonomia oficial")
    return 0


def cmd_discover(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    with SteamClient(settings) as client:
        taxonomy = _load_taxonomy(client, settings)
        _resolve_and_report(taxonomy, filters)
        if args.pilot_pages:
            log.info("[DISCOVERY] MODO PILOTO: %d paginas por tag, gravando em "
                     "candidates_pilot.json (nao afeta a coleta definitiva)",
                     args.pilot_pages)
        discovery = Discovery(client, settings, filters, taxonomy,
                             pilot_pages=args.pilot_pages)
        discovery.run(resume=not args.no_resume, tag_limit=args.tag_limit,
                      build_indie_ledger=not args.skip_indie_ledger
                      and not args.pilot_pages)
    return 0


def _scrape_ids(args: argparse.Namespace, settings, filters) -> list[int]:
    if args.app_ids:
        return [int(x) for x in args.app_ids.split(",") if x.strip()]
    if getattr(args, "prototype", False):
        for app_id, desc in PROTOTYPE_SAMPLE:
            log.info("[SAMPLE]    %-10d %s", app_id, desc)
        return [a for a, _ in PROTOTYPE_SAMPLE]

    pilot = bool(getattr(args, "from_pilot", False))
    ids = load_candidate_ids(settings.path("processed"), pilot=pilot)
    if not ids:
        log.error("candidates%s.json vazio ou ausente. Rode "
                  "`python main.py discover%s` primeiro, ou use "
                  "`--app-ids` / `--prototype`.",
                  "_pilot" if pilot else "",
                  " --pilot-pages N" if pilot else "")
        return []
    if args.start_from:
        try:
            ids = ids[ids.index(int(args.start_from)):]
            log.info("[SCRAPE]    iniciando a partir do app_id %s", args.start_from)
        except ValueError:
            log.warning("app_id %s nao esta na lista de candidatos; "
                        "processando a lista inteira", args.start_from)
    if args.limit:
        ids = ids[:int(args.limit)]
        log.info("[SCRAPE]    limitado a %d apps (modo de teste)", len(ids))
    return ids


def cmd_scrape(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    from steamreq.scraper import Scraper
    ids = _scrape_ids(args, settings, filters)
    if not ids:
        return 1
    with SteamClient(settings) as client:
        scraper = Scraper(client, settings, filters)
        stats = scraper.run(ids, force=args.force)
    log.info("[SCRAPE]    FIM | %s completos | %s parciais | %s falhas | "
             "%s pulados", f"{stats.completed:,}", f"{stats.partial:,}",
             f"{stats.failed:,}", f"{stats.skipped:,}")
    return 0


def cmd_retry_failed(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    from steamreq.scraper import Scraper
    cp = Checkpoint(f"{settings.path('checkpoints')}/scrape_state.json",
                    filter_version=filters.filter_version,
                    scraper_version=settings.scraper_version)
    failed = cp.app_ids_with_status(Status.FAILED)
    if not failed:
        log.info("[RETRY]     nenhuma falha registrada no checkpoint")
        return 0
    log.info("[RETRY]     %s apps com falha serao reprocessados", f"{len(failed):,}")
    n = cp.reset_failed()
    log.info("[RETRY]     %s registros voltaram para PENDING", f"{n:,}")
    with SteamClient(settings) as client:
        scraper = Scraper(client, settings, filters)
        scraper.checkpoint = cp
        stats = scraper.run(failed, force=True)
    log.info("[RETRY]     FIM | %s recuperados | %s ainda com falha",
             f"{stats.completed + stats.partial:,}", f"{stats.failed:,}")
    return 0


def cmd_filter(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    run_filter_stage(settings, filters)
    return 0


def cmd_validate(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    from steamreq.export import build_all_records
    records = build_all_records(settings, filters)
    if not records:
        log.error("nenhum registro em data/raw. Rode `scrape` primeiro.")
        return 1
    report = validate_dataset(records, filters)
    log.info("[VALIDATE]  %s registros verificados | %s erros | %s avisos",
             f"{report.checked:,}", f"{len(report.errors):,}",
             f"{len(report.warnings):,}")
    for code, n in report.counts_by_code().items():
        log.info("[VALIDATE]    %-32s %s", code, f"{n:,}")
    for issue in report.errors[:20]:
        log.error("[VALIDATE]  app %s | %s | %s", issue.app_id, issue.code,
                  issue.detail)
    from steamreq.storage import write_json_atomic
    out = f"{settings.path('processed')}/validation_report.json"
    write_json_atomic(out, report.as_dict(), indent=2)
    log.info("[VALIDATE]  relatorio -> %s", out)
    if report.ok:
        log.info("[VALIDATE]  OK: nenhum erro bloqueante")
    return 0 if report.ok else 2


def cmd_export(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    export_dataset(settings, filters, flat=args.flat,
                   include_excluded=args.include_excluded)
    return 0


def cmd_analyze_thresholds(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    analyze_thresholds(settings, filters)
    return 0


def cmd_audit_bias(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    from steamreq.discovery_audit import MissingAPIKey, run_audit
    try:
        with SteamClient(settings) as client:
            run_audit(client, settings)
    except MissingAPIKey as exc:
        log.warning("[AUDIT]     %s", exc)
        return 0
    _ = filters
    return 0


def cmd_status(args: argparse.Namespace) -> int:
    settings, filters = _bootstrap(args)
    cp = Checkpoint(f"{settings.path('checkpoints')}/scrape_state.json",
                    filter_version=filters.filter_version,
                    scraper_version=settings.scraper_version)
    raw = RawStore(settings.path("raw_appdetails"), settings.path("raw_tags"))
    counts = cp.counts()
    candidates = load_candidate_ids(settings.path("processed"))
    log.info("[STATUS]    candidatos no discovery : %s", f"{len(candidates):,}")
    log.info("[STATUS]    apps no checkpoint      : %s", f"{len(cp.apps):,}")
    for status, n in counts.items():
        if n:
            log.info("[STATUS]      %-20s %s", status, f"{n:,}")
    log.info("[STATUS]    payloads brutos salvos : %s appdetails",
             f"{len(raw.known_app_ids()):,}")
    log.info("[STATUS]    filter_version         : %s (checkpoint: %s)",
             filters.filter_version, cp.meta.get("filter_version"))
    failed = cp.app_ids_with_status(Status.FAILED)
    if failed:
        log.info("[STATUS]    exemplos de falhas     : %s",
                 failed[:10])
        log.info("[STATUS]    use `python main.py retry-failed` para reprocessar")
    return 0


# --- parser -----------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="main.py",
        description="Pipeline de coleta de requisitos de hardware da Steam "
                    "(TCC MBA DSA)",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--settings", default=None, help="caminho de settings.yaml")
    p.add_argument("--filters", default=None, help="caminho de filters.yaml")
    p.add_argument("-v", "--verbose", action="store_true",
                   help="console em DEBUG")
    sub = p.add_subparsers(dest="command", required=True)

    s = sub.add_parser("tags", help="baixa a taxonomia e resolve o filters.yaml")
    s.add_argument("--refresh", action="store_true", help="ignora o cache local")
    s.set_defaults(func=cmd_tags)

    s = sub.add_parser("discover", help="estagio 1: enumera candidatos")
    s.add_argument("--no-resume", action="store_true",
                   help="recomeca do zero, ignorando o estado anterior")
    s.add_argument("--tag-limit", type=int, default=None,
                   help="consulta apenas as N primeiras tags (teste)")
    s.add_argument("--skip-indie-ledger", action="store_true",
                   help="nao coleta o ledger de excluidos por Indie")
    s.add_argument("--pilot-pages", type=int, default=None, metavar="N",
                   help="modo piloto: no maximo N paginas por tag, gravando em "
                        "candidates_pilot.json (nao afeta a coleta definitiva)")
    s.set_defaults(func=cmd_discover)

    for name, helptext in (("scrape", "estagio 2: coleta appdetails + tags"),
                           ("resume", "retoma o scrape de onde parou")):
        s = sub.add_parser(name, help=helptext)
        s.add_argument("--limit", type=int, default=None)
        s.add_argument("--start-from", default=None, metavar="APP_ID")
        s.add_argument("--app-ids", default=None,
                       help="lista explicita separada por virgula")
        s.add_argument("--prototype", action="store_true",
                       help="usa a amostra heterogenea da FASE 3")
        s.add_argument("--from-pilot", action="store_true",
                       help="le os candidatos de candidates_pilot.json")
        s.add_argument("--force", action="store_true",
                       help="reprocessa mesmo o que ja esta concluido")
        s.set_defaults(func=cmd_scrape)

    s = sub.add_parser("retry-failed", help="reprocessa apenas as falhas")
    s.set_defaults(func=cmd_retry_failed)

    s = sub.add_parser("filter", help="estagio 3: aplica a metodologia offline")
    s.set_defaults(func=cmd_filter)

    s = sub.add_parser("validate", help="estagio 5: valida os registros")
    s.set_defaults(func=cmd_validate)

    s = sub.add_parser("export", help="estagio 6: grava dataset.json")
    s.add_argument("--flat", action="store_true",
                   help="achata em minimum_* / recommended_*")
    s.add_argument("--include-excluded", action="store_true",
                   help="inclui tambem os registros excluidos")
    s.set_defaults(func=cmd_export)

    s = sub.add_parser("analyze-thresholds",
                       help="item 2.14: distribuicao de reviews, sem aplicar corte")
    s.set_defaults(func=cmd_analyze_thresholds)

    s = sub.add_parser("audit-bias",
                       help="decisao D1: quantifica o vies de sobrevivencia")
    s.set_defaults(func=cmd_audit_bias)

    s = sub.add_parser("status", help="resumo do checkpoint")
    s.set_defaults(func=cmd_status)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return int(args.func(args) or 0)
    except (ConfigError, TagResolutionError) as exc:
        # Falha ruidosa deliberada: prosseguir produziria filtro invalido.
        print(f"\nERRO DE CONFIGURACAO: {exc}\n", file=sys.stderr)
        return 3
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuario.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
