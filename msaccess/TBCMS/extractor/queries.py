"""Extract query definitions (SQL) from MS Access via DAO."""

import json
import logging
import re
from pathlib import Path

logger = logging.getLogger(__name__)

# System query prefixes to skip
SYSTEM_QUERY_PREFIXES = ("~sq_", "~", "MSys")
QUERY_TYPE_PREFIX_MAP = (
    ("SELECT", "select"),
    ("TRANSFORM", "crosstab"),
    ("INSERT", "insert"),
    ("UPDATE", "update"),
    ("DELETE", "delete"),
    ("CREATE", "ddl"),
    ("ALTER", "ddl"),
    ("DROP", "ddl"),
    ("PARAMETERS", "parameterized"),
    ("EXEC", "pass-through"),
)


def _is_user_query(name: str) -> bool:
    """Return True if query is a user query (not system)."""
    return not any(
        name.startswith(prefix) for prefix in SYSTEM_QUERY_PREFIXES
    )


def _safe_name(name: str) -> str:
    """Sanitize object name for filenames."""
    return "".join(c if c.isalnum() or c in "._-" else "_" for c in name)


def _detect_query_type(sql: str) -> str:
    """Infer broad query type from leading SQL keyword."""
    normalized = sql.lstrip().upper()
    for prefix, query_type in QUERY_TYPE_PREFIX_MAP:
        if normalized.startswith(prefix):
            return query_type
    return "unknown"


def _extract_references(sql: str) -> tuple[list[str], list[str]]:
    """Best-effort extraction of referenced tables and queries."""
    refs = re.findall(
        r"\b(?:FROM|JOIN|INTO|UPDATE)\s+(\[[^\]]+\]|[A-Za-z0-9_.$-]+)",
        sql,
        flags=re.IGNORECASE,
    )
    clean_refs = []
    for ref in refs:
        name = ref.strip("[]")
        if name:
            clean_refs.append(name)

    table_refs: list[str] = []
    query_refs: list[str] = []
    for name in clean_refs:
        lowered = name.lower()
        if lowered.startswith("qry") or lowered.startswith("query"):
            query_refs.append(name)
        else:
            table_refs.append(name)
    return sorted(set(table_refs)), sorted(set(query_refs))


def _extract_parameters(sql: str) -> list[dict]:
    """Best-effort extraction of Access PARAMETERS declaration."""
    match = re.search(r"^\s*PARAMETERS\s+(.*?);", sql, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        return []
    params_blob = match.group(1)
    parts = [p.strip() for p in params_blob.split(",") if p.strip()]
    parameters = []
    for part in parts:
        tokens = part.split()
        if not tokens:
            continue
        param_name = tokens[0].strip("[]")
        data_type = " ".join(tokens[1:]) if len(tokens) > 1 else "UNKNOWN"
        parameters.append({
            "name": param_name,
            "dataType": data_type,
            "inferred": True,
        })
    return parameters


def extract_queries(accdb_path: Path, output_dir: Path) -> list[str]:
    """
    Extract QueryDef SQL from Access database via DAO.
    Returns list of extracted query names.
    """
    try:
        import win32com.client
    except ImportError:
        logger.warning(
            "pywin32 not installed; cannot extract queries. "
            "Install with: pip install pywin32"
        )
        return []

    accdb_path = Path(accdb_path).resolve()
    if not accdb_path.exists():
        logger.error("Database not found: %s", accdb_path)
        return []

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    extracted: list[str] = []
    query_index = {"queries": []}
    dao = None
    db = None

    try:
        # DAO.DBEngine.120 for Access 2016+; fallback to .36 for older
        for prog_id in ("DAO.DBEngine.120", "DAO.DBEngine.36", "DAO.DBEngine.35"):
            try:
                dao = win32com.client.Dispatch(prog_id)
                break
            except Exception:
                continue

        if dao is None:
            logger.warning("DAO.DBEngine not available; cannot extract queries")
            return []

        db = dao.OpenDatabase(str(accdb_path))

        for i in range(db.QueryDefs.Count):
            qdf = db.QueryDefs(i)
            name = qdf.Name
            if not _is_user_query(name):
                continue
            try:
                sql = qdf.SQL.strip()
                if not sql:
                    continue
                safe_name = _safe_name(name)
                out_path = output_dir / f"{safe_name}.sql"
                out_path.write_text(sql, encoding="utf-8")
                extracted.append(name)

                referenced_tables, referenced_queries = _extract_references(sql)
                query_index["queries"].append({
                    "name": name,
                    "sqlFile": f"extract/queries/{safe_name}.sql",
                    "type": _detect_query_type(sql),
                    "parameters": _extract_parameters(sql),
                    "referencedTables": referenced_tables,
                    "referencedQueries": referenced_queries,
                })
            except Exception as e:
                logger.warning("Skipping query %s: %s", name, e)

    except Exception as e:
        logger.error("DAO extraction failed: %s", e)
    finally:
        if db is not None:
            try:
                db.Close()
            except Exception:
                pass

    try:
        index_path = output_dir / "index.json"
        index_path.write_text(
            json.dumps(query_index, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception as e:
        logger.warning("Could not write query index.json: %s", e)

    return extracted
