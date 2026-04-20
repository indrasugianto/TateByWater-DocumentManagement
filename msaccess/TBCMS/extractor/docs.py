"""Generate schema documentation report in Markdown."""

import json
import logging
from pathlib import Path

from .access_app import _safe_filename

logger = logging.getLogger(__name__)


def generate_report(
    output_path: Path,
    schema: dict,
    relationships: list[dict],
    queries: list[str],
    vba: dict[str, list[str]],
    table_counts: dict[str, int] | None = None,
    structured_reports: dict | None = None,
    structured_forms: dict | None = None,
    app_manifest: dict | None = None,
    report_lineage: dict | None = None,
) -> None:
    """
    Generate schema_report.md with tables, relationships, queries, VBA inventory,
    and optionally a structured report inventory.
    """
    lines = []
    # output_path is docs/schema_report.md; app folder is parent.parent
    app_name = output_path.parent.parent.name
    lines.append(f"# Schema Report: {app_name}\n")
    lines.append("---\n")

    # Tables
    lines.append("## Tables\n")
    for table in schema.get("tables", []):
        lines.append(f"### {table['name']}\n")
        lines.append("| Column | Type | Size |\n")
        lines.append("|--------|------|------|\n")
        for col in table.get("columns", []):
            size = col.get("size", "—")
            lines.append(f"| {col['name']} | {col['type']} | {size} |\n")
        if table.get("primary_key"):
            lines.append(f"\n**Primary Key:** {', '.join(table['primary_key'])}\n")
        if table_counts and table["name"] in table_counts:
            lines.append(f"\n**Row count:** {table_counts[table['name']]}\n")
        lines.append("\n")

    # Relationships
    lines.append("## Relationships\n")
    if relationships:
        lines.append("| Table | Column(s) | Referenced Table | Referenced Column(s) |\n")
        lines.append("|-------|-----------|------------------|----------------------|\n")
        for rel in relationships:
            tbl = rel.get("table", "")
            cols = rel.get("columns", [])
            ref_tbl = rel.get("referenced_table", "")
            ref_cols = rel.get("referenced_columns", [])
            if isinstance(cols, str):
                cols = [cols]
            if isinstance(ref_cols, str):
                ref_cols = [ref_cols]
            lines.append(
                f"| {tbl} | {', '.join(cols)} | {ref_tbl} | {', '.join(ref_cols)} |\n"
            )
    else:
        lines.append("No relationships extracted.\n")
    lines.append("\n")

    # Queries
    lines.append("## Queries\n")
    if queries:
        queries_dir = output_path.parent.parent / "extract" / "queries"
        for name in sorted(queries):
            lines.append(f"### {name}\n")
            sql_path = queries_dir / f"{_safe_filename(name)}.sql"
            if sql_path.exists():
                try:
                    sql = sql_path.read_text(encoding="utf-8").strip()
                    sql_preview = sql[:500] + "..." if len(sql) > 500 else sql
                    lines.append("```sql\n")
                    lines.append(sql_preview)
                    lines.append("\n```\n")
                except Exception:
                    lines.append("\n")
            else:
                lines.append("\n")
    else:
        lines.append("No queries extracted.\n")
    lines.append("\n")

    # VBA inventory
    lines.append("## VBA Object Inventory\n")
    lines.append("| Type | Count | Objects |\n")
    lines.append("|------|-------|--------|\n")
    for obj_type in ("forms", "reports", "modules", "classes", "macros"):
        items = vba.get(obj_type, [])
        if items:
            lines.append(f"| {obj_type.capitalize()} | {len(items)} | ")
            lines.append(", ".join(sorted(items)))
            lines.append(" |\n")
        else:
            lines.append(f"| {obj_type.capitalize()} | 0 |  |\n")
    lines.append("\n")

    # Structured Report Inventory
    if structured_reports:
        lines.append("## Structured Report Inventory\n")
        reports_list = structured_reports.get("reports", [])
        errors = structured_reports.get("errors", [])
        lines.append(f"**Extracted:** {len(reports_list)} reports\n\n")
        if reports_list:
            lines.append("| Report | Data Source | Sections | Subreports |\n")
            lines.append("|--------|------------|----------|------------|\n")
            for rpt in reports_list:
                name = rpt.get("reportName", "")
                ds = rpt.get("dataSource", "")
                ds_preview = (ds[:60] + "...") if len(ds) > 60 else ds
                sections = rpt.get("sections", [])
                subreport_count = sum(
                    1
                    for sec in sections
                    for ctrl in sec.get("controls", [])
                    if ctrl.get("type") == "SubReport"
                )
                lines.append(
                    f"| {name} | {ds_preview} | {len(sections)} | {subreport_count} |\n"
                )
            lines.append("\n")
        if errors:
            lines.append("### Extraction Errors\n")
            for err in errors:
                lines.append(f"- **{err.get('report', 'unknown')}**: {err.get('error', '')}\n")
            lines.append("\n")

    # Structured Form Inventory
    if structured_forms:
        lines.append("## Structured Form Inventory\n")
        forms_list = structured_forms.get("forms", [])
        errors = structured_forms.get("errors", [])
        lines.append(f"**Extracted:** {len(forms_list)} forms\n\n")
        if forms_list:
            lines.append("| Form | Record Source | Sections | Controls |\n")
            lines.append("|------|---------------|----------|----------|\n")
            for form in forms_list:
                name = form.get("formName", "")
                rs = form.get("recordSource", "")
                rs_preview = (rs[:60] + "...") if isinstance(rs, str) and len(rs) > 60 else rs
                sections = form.get("sections", [])
                control_count = sum(len(sec.get("controls", [])) for sec in sections)
                lines.append(
                    f"| {name} | {rs_preview} | {len(sections)} | {control_count} |\n"
                )
            lines.append("\n")
        if errors:
            lines.append("### Form Extraction Errors\n")
            for err in errors:
                lines.append(f"- **{err.get('form', 'unknown')}**: {err.get('error', '')}\n")
            lines.append("\n")

    # JSON Layer Summary
    if app_manifest:
        lines.append("## JSON Layer Summary\n")
        coverage = app_manifest.get("coverage", {})
        lines.append("| Artifact | Count |\n")
        lines.append("|----------|-------|\n")
        lines.append(f"| Structured forms (`extract/forms/*.json`) | {coverage.get('formCount', 0)} |\n")
        lines.append(f"| Structured reports (`extract/reports/*.json`) | {coverage.get('reportCount', 0)} |\n")
        lines.append(f"| Query index (`extract/queries/index.json`) | {coverage.get('queryCount', 0)} |\n")
        lines.append(f"| VBA index (`extract/vba/index.json`) | {coverage.get('vbaProcedureCount', 0)} procedures |\n")
        lines.append(f"| Report lineage index (`extract/lineage/index.json`) | {coverage.get('lineageReportCount', 0)} reports |\n")
        lines.append("| App manifest (`extract/app_manifest.json`) | 1 |\n")
        lines.append("\n")

    if report_lineage:
        lines.append("## Report Lineage Summary\n")
        lineage_reports = report_lineage.get("reports", [])
        errors = report_lineage.get("errors", [])
        lines.append(f"**Generated:** {len(lineage_reports)} lineage report(s)\n\n")
        if lineage_reports:
            lines.append("| Report | Trigger Paths | Queries | Tables | Confidence |\n")
            lines.append("|--------|---------------|---------|--------|------------|\n")
            for row in lineage_reports:
                lines.append(
                    f"| {row.get('name', '')} | {row.get('triggerCount', 0)} | "
                    f"{row.get('queryCount', 0)} | {row.get('tableCount', 0)} | "
                    f"{row.get('confidence', '')} |\n"
                )
            lines.append("\n")
        if errors:
            lines.append("### Lineage Errors\n")
            for err in errors:
                lines.append(f"- **{err.get('report', 'unknown')}**: {err.get('error', '')}\n")
            lines.append("\n")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("".join(lines), encoding="utf-8")
