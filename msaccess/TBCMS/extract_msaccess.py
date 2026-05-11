#!/usr/bin/env python3
"""MS Access Extractor - Extract schema, data, queries, VBA, and structured reports from .accdb files."""

import argparse
import logging
import sys
from pathlib import Path

from extractor.schema import extract_schema
from extractor.data import export_table_data
from extractor.relationships import extract_relationships
from extractor.queries import extract_queries
from extractor.vba import extract_vba
from extractor.reports import extract_reports
from extractor.forms import extract_forms
from extractor.lineage import build_report_lineage
from extractor.manifest import build_app_manifest
from extractor.docs import generate_report

logger = logging.getLogger(__name__)

# Default paths relative to project root
MSACCESS_DIR = "msaccess"
ORIGINAL_DIR = "original"
EXTRACT_DIR = "extract"
DOCS_DIR = "docs"
LOCK_ERROR_MARKERS = (
    "prevents it from being opened or locked",
    "placed in a state by",
    "(-3810)",
)


def _get_project_root() -> Path:
    """Project root is parent of code/ directory."""
    return Path(__file__).resolve().parent.parent


def _discover_apps(project_root: Path, app_filter: str | None) -> list[tuple[str, Path]]:
    """Discover (app_name, accdb_path) for each app with original/*.accdb."""
    msaccess = project_root / MSACCESS_DIR
    if not msaccess.is_dir():
        return []

    results = []
    for app_dir in msaccess.iterdir():
        if not app_dir.is_dir():
            continue
        if app_filter and app_dir.name != app_filter:
            continue
        original = app_dir / ORIGINAL_DIR
        if not original.is_dir():
            continue
        for accdb in list(original.glob("*.accdb")) + list(original.glob("*.mdb")):
            results.append((app_dir.name, accdb))
            break  # One db per app
    return results


def _get_connection_string(accdb_path: Path) -> str:
    """Build pyodbc connection string for Access."""
    path = accdb_path.resolve()
    return f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};"


def _get_table_counts(conn_str: str, schema: dict) -> dict[str, int]:
    """Get row counts for each table (best-effort)."""
    try:
        import pyodbc
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        counts = {}
        for table in schema.get("tables", []):
            name = table["name"]
            try:
                cursor.execute(f"SELECT COUNT(*) FROM [{name}]")
                counts[name] = cursor.fetchone()[0]
            except Exception:
                pass
        conn.close()
        return counts
    except Exception:
        return {}


def _is_lock_error(error: Exception | str) -> bool:
    """Return True when Access/ODBC reports exclusive lock state."""
    text = str(error).lower()
    return any(marker in text for marker in LOCK_ERROR_MARKERS)


def extract_app(
    app_name: str,
    accdb_path: Path,
    project_root: Path,
    capture_screenshots: bool = False,
    output_dir: Path | None = None,
    fail_fast_on_lock: bool = True,
    form_timeout_seconds: int = 60,
    allow_direct_dispatch_fallback: bool = False,
) -> bool:
    """Extract all data for one app. Returns True on success."""
    if output_dir is not None:
        app_dir = output_dir
    else:
        app_dir = accdb_path.parent.parent
    extract_path = app_dir / EXTRACT_DIR
    docs_path = app_dir / DOCS_DIR

    extract_path.mkdir(parents=True, exist_ok=True)
    docs_path.mkdir(parents=True, exist_ok=True)

    schema = {"tables": []}
    relationships = []
    queries = []
    vba_result = {"forms": [], "reports": [], "modules": [], "classes": [], "macros": []}
    structured_reports = None
    structured_forms = None
    report_lineage = None
    manifest_data = None

    lock_errors = []
    conn_str = _get_connection_string(accdb_path)

    # pyodbc extraction
    try:
        schema = extract_schema(conn_str, extract_path / "schema.json")
    except Exception as e:
        logger.warning("Schema extraction failed for %s: %s", app_name, e)
        if _is_lock_error(e):
            lock_errors.append(f"schema: {e}")

    try:
        export_table_data(conn_str, extract_path)
    except Exception as e:
        logger.warning("Data export failed for %s: %s", app_name, e)
        if _is_lock_error(e):
            lock_errors.append(f"data: {e}")

    try:
        relationships = extract_relationships(
            conn_str, extract_path / "relationships.json"
        )
    except Exception as e:
        logger.warning("Relationships extraction failed for %s: %s", app_name, e)
        if _is_lock_error(e):
            lock_errors.append(f"relationships: {e}")

    # Query extraction (DAO)
    queries = extract_queries(accdb_path, extract_path / "queries")

    # Fail fast if DB is in locked/exclusive state to avoid COM hangs.
    if lock_errors:
        if fail_fast_on_lock:
            logger.error(
                "Database appears locked in exclusive mode. "
                "Skipping VBA/forms/reports extraction to avoid hangs. "
                "Close all Access sessions using this database and retry. "
                "Use --continue-on-lock to override this safeguard."
            )
            for detail in lock_errors:
                logger.error("  Lock detail: %s", detail)
            return False
        logger.warning(
            "Database lock state detected, but continuing because "
            "--continue-on-lock was specified."
        )
        for detail in lock_errors:
            logger.warning("  Lock detail: %s", detail)

    # VBA extraction (Access.Application)
    vba_result = extract_vba(
        accdb_path,
        extract_path / "vba",
        allow_direct_fallback=allow_direct_dispatch_fallback,
    )

    # Structured form extraction
    forms_path = extract_path / "forms"
    vba_forms_dir = extract_path / "vba" / "forms"
    structured_forms = extract_forms(
        accdb_path, forms_path,
        form_names=vba_result.get("forms", []),
        vba_dir=vba_forms_dir,
        per_form_timeout_seconds=form_timeout_seconds,
        allow_direct_fallback=allow_direct_dispatch_fallback,
    )
    logger.info(
        "  Forms: %d extracted, %d errors",
        structured_forms["count"],
        len(structured_forms["errors"]),
    )

    # Structured report extraction
    reports_path = extract_path / "reports"
    vba_reports_dir = extract_path / "vba" / "reports"
    structured_reports = extract_reports(
        accdb_path, reports_path,
        capture_screenshots=capture_screenshots,
        vba_dir=vba_reports_dir,
        allow_direct_fallback=allow_direct_dispatch_fallback,
    )
    logger.info(
        "  Reports: %d extracted, %d errors",
        structured_reports["count"],
        len(structured_reports["errors"]),
    )

    # Report lineage extraction (best effort) from structured artifacts.
    try:
        report_lineage = build_report_lineage(
            extract_dir=extract_path,
            output_dir=extract_path / "lineage",
            structured_reports=structured_reports,
            structured_forms=structured_forms,
        )
        logger.info(
            "  Lineage: %d reports, %d errors",
            report_lineage.get("count", 0),
            len(report_lineage.get("errors", [])),
        )
    except Exception as e:
        logger.warning("Lineage generation failed for %s: %s", app_name, e)

    # Table counts for report
    table_counts = _get_table_counts(conn_str, schema)

    # Build app manifest using available extraction outputs
    try:
        manifest_data = build_app_manifest(
            app_name=app_name,
            source_file=accdb_path.name,
            extract_dir=extract_path,
            output_path=extract_path / "app_manifest.json",
            schema=schema,
            relationships=relationships,
            queries=queries,
            vba_result=vba_result,
            structured_reports=structured_reports,
            structured_forms=structured_forms,
        )
    except Exception as e:
        logger.warning("App manifest generation failed for %s: %s", app_name, e)

    # Generate docs
    generate_report(
        docs_path / "schema_report.md",
        schema=schema,
        relationships=relationships,
        queries=queries,
        vba=vba_result,
        table_counts=table_counts,
        structured_reports=structured_reports,
        structured_forms=structured_forms,
        app_manifest=manifest_data,
        report_lineage=report_lineage,
    )

    return True


def main() -> int:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Extract schema, data, queries, VBA, and structured reports from MS Access databases."
    )
    parser.add_argument(
        "--source",
        metavar="FILE",
        help="Path to a specific .accdb or .mdb file to extract. "
             "When used with --output-dir, bypasses the msaccess/ folder discovery.",
    )
    parser.add_argument(
        "--app",
        metavar="NAME",
        help="Process only this app (subfolder name under msaccess/)",
    )
    parser.add_argument(
        "--screenshots",
        action="store_true",
        help="Capture designer and preview screenshots for reports.",
    )
    parser.add_argument(
        "--output-dir",
        metavar="DIR",
        help="Root output directory. Extracted files go into DIR/extract/ and DIR/docs/. "
             "If omitted, output goes next to the original .accdb file.",
    )
    parser.add_argument(
        "--continue-on-lock",
        action="store_true",
        help="Do not fail fast when database lock state is detected. "
             "Attempts to continue VBA/forms/reports extraction (may hang if Access is locked).",
    )
    parser.add_argument(
        "--allow-direct-dispatch-fallback",
        action="store_true",
        help="Allow fallback to direct COM OpenCurrentDatabase if subprocess /nostartup "
             "connection fails. Disabled by default to avoid startup macro/form execution.",
    )
    parser.add_argument(
        "--form-timeout-seconds",
        type=int,
        default=60,
        help="Timeout in seconds for extracting a single structured form. "
             "Forms that exceed this timeout are skipped with an error. Default: 60.",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Verbose logging",
    )
    args = parser.parse_args()

    if args.form_timeout_seconds <= 0:
        parser.error("--form-timeout-seconds must be greater than 0")

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
        stream=sys.stderr,
    )
    if not args.allow_direct_dispatch_fallback:
        logger.info(
            "Startup suppression mode enabled: using Access subprocess with /nostartup only."
        )

    resolved_output_dir = Path(args.output_dir).resolve() if args.output_dir else None

    # --source mode: extract a single file directly (no folder discovery)
    if args.source:
        source_path = Path(args.source).resolve()
        if not source_path.is_file():
            logger.error("Source file not found: %s", source_path)
            return 1
        if resolved_output_dir is None:
            parser.error("--output-dir is required when using --source")

        app_name = source_path.stem
        logger.info("Extracting %s from %s", app_name, source_path)
        logger.info("Output directory: %s", resolved_output_dir)

        if extract_app(
            app_name,
            source_path,
            source_path.parent,
            capture_screenshots=args.screenshots,
            output_dir=resolved_output_dir,
            fail_fast_on_lock=not args.continue_on_lock,
            form_timeout_seconds=args.form_timeout_seconds,
            allow_direct_dispatch_fallback=args.allow_direct_dispatch_fallback,
        ):
            logger.info("  Done: %s/extract/ and %s/docs/", resolved_output_dir, resolved_output_dir)
            return 0
        else:
            logger.error("  Failed: %s", app_name)
            return 1

    # Discovery mode: scan msaccess/ folder structure
    project_root = _get_project_root()
    apps = _discover_apps(project_root, args.app)

    if not apps:
        logger.error(
            "No .accdb files found under %s/<app>/%s/. "
            "Use --app NAME to target a specific app, or --source FILE to extract a specific file.",
            MSACCESS_DIR, ORIGINAL_DIR,
        )
        return 1

    success_count = 0
    for app_name, accdb_path in apps:
        logger.info("Extracting %s (%s)", app_name, accdb_path.name)
        if extract_app(
            app_name,
            accdb_path,
            project_root,
            capture_screenshots=args.screenshots,
            output_dir=resolved_output_dir,
            fail_fast_on_lock=not args.continue_on_lock,
            form_timeout_seconds=args.form_timeout_seconds,
            allow_direct_dispatch_fallback=args.allow_direct_dispatch_fallback,
        ):
            success_count += 1
            out_label = resolved_output_dir or f"{app_name}"
            logger.info("  Done: %s/extract/ and %s/docs/", out_label, out_label)
        else:
            logger.error("  Failed: %s", app_name)

    return 0 if success_count == len(apps) else 1


if __name__ == "__main__":
    sys.exit(main())
