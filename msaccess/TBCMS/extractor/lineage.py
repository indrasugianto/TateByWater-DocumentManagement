"""Build report lineage artifacts from extracted forms/reports/queries/VBA."""

import json
import logging
import re
from pathlib import Path

from .access_app import _safe_filename

logger = logging.getLogger(__name__)


_QUERY_REF_RE = re.compile(
    r"\b(?:FROM|JOIN|INTO|UPDATE)\s+(\[[^\]]+\]|[A-Za-z0-9_.$-]+)",
    flags=re.IGNORECASE,
)

_OPEN_REPORT_CALL_RE = re.compile(r"\bDoCmd\.OpenReport\b", flags=re.IGNORECASE)
_QUOTED_NAME_RE = re.compile(r'"([^"]+)"')

_EVENT_SUFFIX_MAP = {
    "onClick": "Click",
    "onDblClick": "DblClick",
    "onChange": "Change",
    "onBeforeUpdate": "BeforeUpdate",
    "onAfterUpdate": "AfterUpdate",
    "onOpen": "Open",
    "onLoad": "Load",
    "onCurrent": "Current",
    "onClose": "Close",
}


def _safe_read_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _safe_read_text(path: Path) -> str:
    for encoding in ("utf-16-le", "utf-8", "latin-1"):
        try:
            return path.read_text(encoding=encoding, errors="replace")
        except Exception:
            continue
    return ""


def _extract_sql_refs(sql: str) -> tuple[list[str], list[str]]:
    refs = []
    for m in _QUERY_REF_RE.findall(sql or ""):
        name = m.strip("[]").strip()
        if name:
            refs.append(name)
    tables: list[str] = []
    queries: list[str] = []
    for name in refs:
        if name.lower().startswith(("qry", "query")):
            queries.append(name)
        else:
            tables.append(name)
    return sorted(set(tables)), sorted(set(queries))


def _build_query_lookup(queries_index: dict) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for q in queries_index.get("queries", []):
        name = q.get("name")
        if isinstance(name, str) and name:
            out[name.lower()] = q
    return out


def _expand_query_dependencies(start_query: str, queries_lookup: dict[str, dict]) -> dict:
    queue = [start_query]
    seen: set[str] = set()
    involved_queries: list[str] = []
    terminal_tables: set[str] = set()
    edges: list[dict] = []

    while queue:
        current = queue.pop(0)
        key = current.lower()
        if key in seen:
            continue
        seen.add(key)
        involved_queries.append(current)

        node = queries_lookup.get(key, {})
        child_queries = node.get("referencedQueries", []) or []
        child_tables = node.get("referencedTables", []) or []

        for cq in child_queries:
            if not isinstance(cq, str) or not cq:
                continue
            edges.append({"from": current, "to": cq, "type": "query"})
            queue.append(cq)
        for ct in child_tables:
            if not isinstance(ct, str) or not ct:
                continue
            edges.append({"from": current, "to": ct, "type": "table"})
            terminal_tables.add(ct)

    return {
        "involvedQueries": involved_queries,
        "terminalTables": sorted(terminal_tables),
        "queryEdges": edges,
    }


def _candidate_proc_names(form_name: str, control_name: str | None, event_name: str) -> list[str]:
    suffix = _EVENT_SUFFIX_MAP.get(event_name, "")
    if not suffix:
        return []
    names = []
    if control_name:
        names.append(f"{control_name}_{suffix}")
    if event_name in {"onOpen", "onLoad", "onCurrent", "onClose"}:
        names.append(f"Form_{suffix}")
    names.append(f"{form_name}_{suffix}")
    # preserve order while deduping
    deduped: list[str] = []
    seen: set[str] = set()
    for n in names:
        if n.lower() in seen:
            continue
        seen.add(n.lower())
        deduped.append(n)
    return deduped


def _extract_proc_segment(source_path: Path, line_start: int, line_end: int) -> str:
    content = _safe_read_text(source_path)
    if not content:
        return ""
    lines = content.splitlines()
    start_idx = max(line_start - 1, 0)
    end_idx = min(line_end, len(lines))
    if start_idx >= end_idx:
        return ""
    return "\n".join(lines[start_idx:end_idx])


def _find_report_open_hits(segment: str, report_name: str) -> dict | None:
    if not segment or not _OPEN_REPORT_CALL_RE.search(segment):
        return None
    report_lower = report_name.lower()
    quoted = [q.strip() for q in _QUOTED_NAME_RE.findall(segment)]
    direct_match = any(q.lower() == report_lower for q in quoted)
    mention_match = report_lower in segment.lower()
    if not direct_match and not mention_match:
        return None
    confidence = "high" if direct_match else "medium"
    return {
        "reportMentionType": "literal" if direct_match else "inferred",
        "confidence": confidence,
    }


def _collect_trigger_paths(
    report_name: str,
    forms: list[dict],
    vba_procedures: list[dict],
    extract_dir: Path,
) -> list[dict]:
    out: list[dict] = []
    form_proc_lookup: dict[tuple[str, str], dict] = {}
    for proc in vba_procedures:
        if proc.get("objectType") != "form":
            continue
        form_name = proc.get("objectName")
        proc_name = proc.get("procedureName")
        if not isinstance(form_name, str) or not isinstance(proc_name, str):
            continue
        form_proc_lookup[(form_name.lower(), proc_name.lower())] = proc

    for form in forms:
        form_name = form.get("formName")
        if not isinstance(form_name, str) or not form_name:
            continue

        # Form-level events
        for event_name, event_value in (form.get("formEvents", {}) or {}).items():
            if not isinstance(event_value, str) or not event_value:
                continue
            candidates = _candidate_proc_names(form_name, None, event_name)
            for proc_name in candidates:
                proc = form_proc_lookup.get((form_name.lower(), proc_name.lower()))
                if not proc:
                    continue
                source_rel = proc.get("sourceFile") or ""
                source_path = extract_dir / Path(str(source_rel).replace("extract/", ""))
                segment = _extract_proc_segment(
                    source_path,
                    int(proc.get("lineStart", 1)),
                    int(proc.get("lineEnd", 1)),
                )
                hit = _find_report_open_hits(segment, report_name)
                if not hit:
                    continue
                out.append({
                    "formName": form_name,
                    "controlName": None,
                    "eventName": event_name,
                    "procedureName": proc_name,
                    "sourceFile": source_rel,
                    "lineStart": int(proc.get("lineStart", 1)),
                    "lineEnd": int(proc.get("lineEnd", 1)),
                    "confidence": hit["confidence"],
                    "reportMentionType": hit["reportMentionType"],
                })

        # Control-level events
        for section in form.get("sections", []) or []:
            for control in section.get("controls", []) or []:
                control_name = control.get("name")
                events = control.get("events", {}) or {}
                if not isinstance(control_name, str) or not isinstance(events, dict):
                    continue
                for event_name, event_value in events.items():
                    if not isinstance(event_value, str) or not event_value:
                        continue
                    candidates = _candidate_proc_names(form_name, control_name, event_name)
                    for proc_name in candidates:
                        proc = form_proc_lookup.get((form_name.lower(), proc_name.lower()))
                        if not proc:
                            continue
                        source_rel = proc.get("sourceFile") or ""
                        source_path = extract_dir / Path(str(source_rel).replace("extract/", ""))
                        segment = _extract_proc_segment(
                            source_path,
                            int(proc.get("lineStart", 1)),
                            int(proc.get("lineEnd", 1)),
                        )
                        hit = _find_report_open_hits(segment, report_name)
                        if not hit:
                            continue
                        out.append({
                            "formName": form_name,
                            "controlName": control_name,
                            "eventName": event_name,
                            "procedureName": proc_name,
                            "sourceFile": source_rel,
                            "lineStart": int(proc.get("lineStart", 1)),
                            "lineEnd": int(proc.get("lineEnd", 1)),
                            "confidence": hit["confidence"],
                            "reportMentionType": hit["reportMentionType"],
                        })

    # Deduplicate by core key
    deduped: list[dict] = []
    seen: set[tuple] = set()
    for row in out:
        key = (
            row.get("formName"),
            row.get("controlName"),
            row.get("eventName"),
            row.get("procedureName"),
            row.get("lineStart"),
            row.get("lineEnd"),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)
    return deduped


def _derive_data_lineage(report: dict, queries_lookup: dict[str, dict]) -> dict:
    data_source = report.get("dataSource")
    if not isinstance(data_source, str) or not data_source.strip():
        return {
            "recordSource": None,
            "recordSourceType": "unknown",
            "involvedQueries": [],
            "terminalTables": [],
            "queryEdges": [],
        }
    record_source = data_source.strip()
    query_node = queries_lookup.get(record_source.lower())
    if query_node:
        expanded = _expand_query_dependencies(record_source, queries_lookup)
        return {
            "recordSource": record_source,
            "recordSourceType": "saved-query",
            "involvedQueries": expanded["involvedQueries"],
            "terminalTables": expanded["terminalTables"],
            "queryEdges": expanded["queryEdges"],
        }

    if re.match(r"^\s*(SELECT|TRANSFORM|INSERT|UPDATE|DELETE|PARAMETERS)\b", record_source, flags=re.IGNORECASE):
        tables, queries = _extract_sql_refs(record_source)
        expanded_edges: list[dict] = []
        expanded_queries: list[str] = list(queries)
        expanded_tables: set[str] = set(tables)
        for query_name in queries:
            if query_name.lower() in queries_lookup:
                expanded = _expand_query_dependencies(query_name, queries_lookup)
                expanded_queries.extend(expanded["involvedQueries"])
                expanded_tables.update(expanded["terminalTables"])
                expanded_edges.extend(expanded["queryEdges"])
        return {
            "recordSource": record_source,
            "recordSourceType": "inline-sql",
            "involvedQueries": sorted(set(expanded_queries)),
            "terminalTables": sorted(expanded_tables),
            "queryEdges": expanded_edges,
        }

    return {
        "recordSource": record_source,
        "recordSourceType": "table-or-unknown",
        "involvedQueries": [],
        "terminalTables": [record_source],
        "queryEdges": [],
    }


def _collect_related_vba_procs(
    report_name: str,
    data_lineage: dict,
    vba_procedures: list[dict],
    extract_dir: Path,
) -> list[dict]:
    related: list[dict] = []
    tokens = {report_name.lower()}
    for q in data_lineage.get("involvedQueries", []):
        if isinstance(q, str):
            tokens.add(q.lower())
    for t in data_lineage.get("terminalTables", []):
        if isinstance(t, str):
            tokens.add(t.lower())

    for proc in vba_procedures:
        source_rel = proc.get("sourceFile")
        if not isinstance(source_rel, str) or not source_rel:
            continue
        source_path = extract_dir / Path(source_rel.replace("extract/", ""))
        segment = _extract_proc_segment(
            source_path,
            int(proc.get("lineStart", 1)),
            int(proc.get("lineEnd", 1)),
        )
        if not segment:
            continue
        lowered = segment.lower()
        if not any(token in lowered for token in tokens):
            continue
        related.append({
            "objectType": proc.get("objectType"),
            "objectName": proc.get("objectName"),
            "procedureName": proc.get("procedureName"),
            "sourceFile": source_rel,
            "lineStart": int(proc.get("lineStart", 1)),
            "lineEnd": int(proc.get("lineEnd", 1)),
            "usesDoCmd": proc.get("usesDoCmd", []),
            "usesSql": bool(proc.get("usesSql", False)),
        })

    return related


def _to_markdown(lineage: dict) -> str:
    report_name = lineage.get("reportName", "UnknownReport")
    lines = [
        f"# Report Lineage: {report_name}",
        "",
        "## Trigger Paths",
    ]
    triggers = lineage.get("triggerPaths", [])
    if triggers:
        for t in triggers:
            control = t.get("controlName") or "(form-level)"
            lines.append(
                f"- {t.get('formName')} -> {control} -> {t.get('eventName')} -> "
                f"{t.get('procedureName')} ({t.get('confidence')} confidence)"
            )
    else:
        lines.append("- No trigger path could be inferred from extracted forms/VBA.")

    lines.extend(["", "## Data Lineage"])
    data_lineage = lineage.get("dataLineage", {})
    lines.append(f"- RecordSource: `{data_lineage.get('recordSource')}`")
    lines.append(f"- RecordSourceType: `{data_lineage.get('recordSourceType')}`")
    iq = data_lineage.get("involvedQueries", [])
    tt = data_lineage.get("terminalTables", [])
    lines.append(f"- Involved Queries: {', '.join(iq) if iq else '(none)'}")
    lines.append(f"- Terminal Tables: {', '.join(tt) if tt else '(none)'}")

    lines.extend(["", "## Related VBA Procedures"])
    related = lineage.get("relatedVbaProcedures", [])
    if related:
        for proc in related:
            lines.append(
                f"- {proc.get('objectType')} {proc.get('objectName')}::{proc.get('procedureName')} "
                f"[{proc.get('lineStart')}-{proc.get('lineEnd')}]"
            )
    else:
        lines.append("- No related VBA procedure could be inferred.")

    return "\n".join(lines).strip() + "\n"


def build_report_lineage(
    extract_dir: Path,
    output_dir: Path,
    structured_reports: dict | None = None,
    structured_forms: dict | None = None,
) -> dict:
    """Generate per-report lineage JSON and Markdown files."""
    extract_dir = Path(extract_dir)
    output_dir = Path(output_dir)
    reports_dir = output_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    reports = (structured_reports or {}).get("reports", [])
    forms = (structured_forms or {}).get("forms", [])
    queries_index = _safe_read_json(extract_dir / "queries" / "index.json", {"queries": []})
    vba_index = _safe_read_json(extract_dir / "vba" / "index.json", {"procedures": [], "globals": []})
    queries_lookup = _build_query_lookup(queries_index)
    vba_procedures = vba_index.get("procedures", [])

    out = {"reports": [], "errors": [], "count": 0}
    for report in reports:
        report_name = report.get("reportName")
        if not isinstance(report_name, str) or not report_name:
            continue
        try:
            trigger_paths = _collect_trigger_paths(report_name, forms, vba_procedures, extract_dir)
            data_lineage = _derive_data_lineage(report, queries_lookup)
            related_vba = _collect_related_vba_procs(
                report_name, data_lineage, vba_procedures, extract_dir
            )

            lineage = {
                "reportName": report_name,
                "dataSource": report.get("dataSource"),
                "triggerPaths": trigger_paths,
                "dataLineage": data_lineage,
                "relatedVbaProcedures": related_vba,
                "confidence": "high" if trigger_paths else "medium",
            }

            safe_name = _safe_filename(report_name)
            json_path = reports_dir / f"{safe_name}.json"
            md_path = reports_dir / f"{safe_name}.md"
            json_path.write_text(json.dumps(lineage, indent=2, ensure_ascii=False), encoding="utf-8")
            md_path.write_text(_to_markdown(lineage), encoding="utf-8")

            out["reports"].append({
                "name": report_name,
                "jsonFile": f"extract/lineage/reports/{safe_name}.json",
                "markdownFile": f"extract/lineage/reports/{safe_name}.md",
                "triggerCount": len(trigger_paths),
                "queryCount": len(data_lineage.get("involvedQueries", [])),
                "tableCount": len(data_lineage.get("terminalTables", [])),
                "confidence": lineage["confidence"],
            })
        except Exception as e:
            logger.warning("Lineage build failed for report %s: %s", report_name, e)
            out["errors"].append({"report": report_name, "error": str(e)})

    out["count"] = len(out["reports"])
    try:
        index_path = output_dir / "index.json"
        index_path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        logger.warning("Could not write lineage index.json: %s", e)

    return out
