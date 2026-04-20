"""Build a normalized app_manifest.json from extracted artifacts."""

import json
from datetime import datetime, timezone
from pathlib import Path


def _safe_read_json(path: Path, default):
    """Read JSON with fallback to default on error."""
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _safe_read_text(path: Path, default: str = "") -> str:
    """Read text file with UTF-8 fallback behavior."""
    try:
        if path.exists():
            return path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        pass
    return default


def _extract_feature_candidates(forms: list[dict], reports: list[dict]) -> list[dict]:
    """Generate best-effort feature seeds from forms and reports."""
    candidates = []
    for idx, form in enumerate(forms, start=1):
        name = form.get("formName", f"UnknownForm{idx}")
        deps = form.get("dependencies", {})
        candidates.append({
            "featureId": f"F{idx:03d}",
            "name": name,
            "primaryForms": [name],
            "supportingQueries": deps.get("queries", []),
            "supportingReports": [],
            "riskLevel": "medium",
            "heuristic": True,
        })

    report_start = len(candidates) + 1
    for i, report in enumerate(reports, start=report_start):
        name = report.get("reportName", f"UnknownReport{i}")
        candidates.append({
            "featureId": f"F{i:03d}",
            "name": f"Reporting - {name}",
            "primaryForms": [],
            "supportingQueries": [report.get("dataSource")] if report.get("dataSource") else [],
            "supportingReports": [name],
            "riskLevel": "medium",
            "heuristic": True,
        })
    return candidates


def build_app_manifest(
    app_name: str,
    source_file: str,
    extract_dir: Path,
    output_path: Path,
    schema: dict,
    relationships: list[dict],
    queries: list[str],
    vba_result: dict,
    structured_reports: dict | None = None,
    structured_forms: dict | None = None,
) -> dict:
    """Build and write app_manifest.json from available extraction outputs."""
    extract_dir = Path(extract_dir)
    queries_index = _safe_read_json(extract_dir / "queries" / "index.json", {"queries": []})
    vba_index = _safe_read_json(extract_dir / "vba" / "index.json", {"procedures": [], "globals": []})
    lineage_index = _safe_read_json(extract_dir / "lineage" / "index.json", {"reports": []})

    forms = (structured_forms or {}).get("forms", [])
    reports = (structured_reports or {}).get("reports", [])

    query_nodes = queries_index.get("queries", [])
    query_names = [q.get("name") for q in query_nodes if q.get("name")]
    if not query_nodes and queries:
        for query_name in queries:
            safe_name = "".join(c if c.isalnum() or c in "._-" else "_" for c in query_name)
            sql_file = extract_dir / "queries" / f"{safe_name}.sql"
            query_nodes.append({
                "name": query_name,
                "sqlFile": f"extract/queries/{safe_name}.sql",
                "type": "unknown",
                "referencedTables": [],
                "referencedQueries": [],
                "sqlPreview": _safe_read_text(sql_file)[:250],
            })
        query_names = queries

    object_forms = [
        {
            "name": f.get("formName"),
            "recordSource": f.get("recordSource"),
            "eventsCount": len(f.get("formEvents", {})),
            "controlsCount": sum(len(s.get("controls", [])) for s in f.get("sections", [])),
            "dependencies": f.get("dependencies", {}),
        }
        for f in forms
    ]

    object_reports = [
        {
            "name": r.get("reportName"),
            "dataSource": r.get("dataSource"),
            "subreports": [
                c.get("sourceObject", "")
                for sec in r.get("sections", [])
                for c in sec.get("controls", [])
                if c.get("type") == "SubReport"
            ],
            "hasVbaCodeBehind": bool(r.get("vbaCodeBehind")),
        }
        for r in reports
    ]

    manifest = {
        "appName": app_name,
        "sourceFile": source_file,
        "extractedAtUtc": datetime.now(timezone.utc).isoformat(),
        "objects": {
            "forms": object_forms,
            "reports": object_reports,
            "queries": query_nodes,
            "modules": [{"name": n} for n in vba_result.get("modules", [])],
            "classes": [{"name": n} for n in vba_result.get("classes", [])],
            "macros": [{"name": n} for n in vba_result.get("macros", [])],
            "tables": [t.get("name") for t in schema.get("tables", [])],
            "relationships": relationships,
            "reportLineage": lineage_index.get("reports", []),
        },
        "featureCandidates": _extract_feature_candidates(forms, reports),
        "migrationHints": {
            "startupObject": object_forms[0]["name"] if object_forms else None,
            "globalStatePatterns": [],
            "securityPatterns": [],
            "highRiskAreas": [
                "dynamic SQL and event-driven VBA behavior may require manual rewrite validation",
                "query dependencies inferred best-effort from SQL text",
            ],
            "heuristic": True,
        },
        "coverage": {
            "formCount": len(forms),
            "reportCount": len(reports),
            "queryCount": len(query_names),
            "vbaProcedureCount": len(vba_index.get("procedures", [])),
            "tableCount": len(schema.get("tables", [])),
            "lineageReportCount": len(lineage_index.get("reports", [])),
        },
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return manifest
