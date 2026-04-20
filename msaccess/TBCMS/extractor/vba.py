"""Extract VBA code from forms, reports, modules, and class modules via Access.Application.SaveAsText."""

import json
import logging
import re
from pathlib import Path

from .access_app import _get_access_app, _quit_access, _safe_filename

logger = logging.getLogger(__name__)

# acForm=2, acReport=3, acModule=5
AC_FORM = 2
AC_REPORT = 3
AC_MODULE = 5

# Class module content hint
CLASS_MODULE_PREFIX = "VERSION 1.0 CLASS"
PROCEDURE_RE = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
    r"(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"([A-Za-z_][A-Za-z0-9_]*)",
    flags=re.IGNORECASE | re.MULTILINE,
)
CALL_RE = re.compile(r"\bCall\s+([A-Za-z_][A-Za-z0-9_]*)", flags=re.IGNORECASE)
DOCMD_RE = re.compile(r"\bDoCmd\.([A-Za-z_][A-Za-z0-9_]*)", flags=re.IGNORECASE)
DECL_RE = re.compile(
    r"^\s*(Public|Global)\s+(?:Const\s+)?([A-Za-z_][A-Za-z0-9_]*)",
    flags=re.IGNORECASE | re.MULTILINE,
)


def _read_text_with_fallback(path: Path) -> str:
    """Read SaveAsText output with encoding fallbacks."""
    for encoding in ("utf-16-le", "utf-8", "latin-1"):
        try:
            return path.read_text(encoding=encoding, errors="replace")
        except Exception:
            continue
    return ""


def _extract_vba_index_entry(
    object_type: str, object_name: str, source_file: str, content: str
) -> dict:
    """Create a best-effort symbol/call index entry for one VBA artifact."""
    lines = content.splitlines()
    line_lookup = {}
    for idx, line in enumerate(lines, start=1):
        line_lookup[idx] = line

    procedures = []
    matches = list(PROCEDURE_RE.finditer(content))
    for i, match in enumerate(matches):
        kind = match.group(1)
        proc_name = match.group(2)

        start_line = content[:match.start()].count("\n") + 1
        if i + 1 < len(matches):
            end_line = content[:matches[i + 1].start()].count("\n")
        else:
            end_line = len(lines)

        segment_lines = [
            line_lookup.get(line_no, "")
            for line_no in range(start_line, min(end_line, len(lines)) + 1)
        ]
        segment = "\n".join(segment_lines)

        calls = sorted(set(CALL_RE.findall(segment)))
        uses_docmd = sorted(set(DOCMD_RE.findall(segment)))
        uses_sql = any(token in segment.upper() for token in ("SELECT ", "INSERT ", "UPDATE ", "DELETE "))

        procedures.append({
            "objectType": object_type,
            "objectName": object_name,
            "moduleName": object_name,
            "procedureName": proc_name,
            "kind": kind.lower(),
            "lineStart": start_line,
            "lineEnd": end_line,
            "calls": calls,
            "usesDoCmd": uses_docmd,
            "usesSql": uses_sql,
            "sourceFile": source_file,
        })

    globals_found = []
    for decl in DECL_RE.finditer(content):
        globals_found.append({
            "moduleName": object_name,
            "name": decl.group(2),
            "declarationScope": decl.group(1).lower(),
            "sourceFile": source_file,
        })

    return {"procedures": procedures, "globals": globals_found}


def extract_vba(
    accdb_path: Path,
    output_dir: Path,
    allow_direct_fallback: bool = False,
) -> dict[str, list[str]]:
    """
    Extract VBA from forms, reports, modules, and class modules via SaveAsText.
    Returns dict with keys: forms, reports, modules, classes (each a list of names).
    """
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        logger.warning(
            "pywin32 not installed; cannot extract VBA. "
            "Install with: pip install pywin32"
        )
        return {"forms": [], "reports": [], "modules": [], "classes": [], "macros": []}

    accdb_path = Path(accdb_path).resolve()
    if not accdb_path.exists():
        logger.error("Database not found: %s", accdb_path)
        return {"forms": [], "reports": [], "modules": [], "classes": [], "macros": []}

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    result = {"forms": [], "reports": [], "modules": [], "classes": [], "macros": []}
    vba_index = {"procedures": [], "globals": []}

    app = None
    try:
        app = _get_access_app(
            str(accdb_path),
            allow_direct_fallback=allow_direct_fallback,
        )

        # Forms
        forms_dir = output_dir / "forms"
        forms_dir.mkdir(parents=True, exist_ok=True)
        form_count = app.CurrentProject.AllForms.Count
        logger.info("Exporting %d forms...", form_count)
        for i in range(form_count):
            obj = app.CurrentProject.AllForms.Item(i)
            name = obj.Name
            try:
                out_path = forms_dir / f"{_safe_filename(name)}.txt"
                app.SaveAsText(AC_FORM, name, str(out_path))
                result["forms"].append(name)
                content = _read_text_with_fallback(out_path)
                source_path = f"extract/vba/forms/{out_path.name}"
                index_entry = _extract_vba_index_entry("form", name, source_path, content)
                vba_index["procedures"].extend(index_entry["procedures"])
                vba_index["globals"].extend(index_entry["globals"])
            except Exception as e:
                logger.warning("Skipping form %s: %s", name, e)

        # Reports
        reports_dir = output_dir / "reports"
        reports_dir.mkdir(parents=True, exist_ok=True)
        report_count = app.CurrentProject.AllReports.Count
        logger.info("Exporting %d reports...", report_count)
        for i in range(report_count):
            obj = app.CurrentProject.AllReports.Item(i)
            name = obj.Name
            try:
                out_path = reports_dir / f"{_safe_filename(name)}.txt"
                app.SaveAsText(AC_REPORT, name, str(out_path))
                result["reports"].append(name)
                content = _read_text_with_fallback(out_path)
                source_path = f"extract/vba/reports/{out_path.name}"
                index_entry = _extract_vba_index_entry("report", name, source_path, content)
                vba_index["procedures"].extend(index_entry["procedures"])
                vba_index["globals"].extend(index_entry["globals"])
            except Exception as e:
                logger.warning("Skipping report %s: %s", name, e)

        # Modules and class modules (AllModules contains both)
        modules_dir = output_dir / "modules"
        classes_dir = output_dir / "classes"
        modules_dir.mkdir(parents=True, exist_ok=True)
        classes_dir.mkdir(parents=True, exist_ok=True)

        module_count = app.CurrentProject.AllModules.Count
        logger.info("Exporting %d modules...", module_count)
        for i in range(module_count):
            obj = app.CurrentProject.AllModules.Item(i)
            name = obj.Name
            try:
                tmp_path = output_dir / "_tmp_module.txt"
                app.SaveAsText(AC_MODULE, name, str(tmp_path))
                content = tmp_path.read_text(encoding="utf-8", errors="replace")
                tmp_path.unlink(missing_ok=True)

                if content.strip().upper().startswith(CLASS_MODULE_PREFIX):
                    out_path = classes_dir / f"{_safe_filename(name)}.txt"
                    result["classes"].append(name)
                    object_type = "class"
                else:
                    out_path = modules_dir / f"{_safe_filename(name)}.txt"
                    result["modules"].append(name)
                    object_type = "module"
                out_path.write_text(content, encoding="utf-8")
                source_path = f"extract/vba/{'classes' if object_type == 'class' else 'modules'}/{out_path.name}"
                index_entry = _extract_vba_index_entry(object_type, name, source_path, content)
                vba_index["procedures"].extend(index_entry["procedures"])
                vba_index["globals"].extend(index_entry["globals"])
            except Exception as e:
                logger.warning("Skipping module %s: %s", name, e)

        # Macros (acMacro=4) — export as text for completeness
        macros_dir = output_dir / "macros"
        macros_dir.mkdir(parents=True, exist_ok=True)
        try:
            macro_count = app.CurrentProject.AllMacros.Count
            if macro_count > 0:
                logger.info("Exporting %d macros...", macro_count)
                for i in range(macro_count):
                    obj = app.CurrentProject.AllMacros.Item(i)
                    name = obj.Name
                    try:
                        out_path = macros_dir / f"{_safe_filename(name)}.txt"
                        app.SaveAsText(4, name, str(out_path))  # acMacro=4
                        result["macros"].append(name)
                        content = _read_text_with_fallback(out_path)
                        source_path = f"extract/vba/macros/{out_path.name}"
                        index_entry = _extract_vba_index_entry("macro", name, source_path, content)
                        vba_index["procedures"].extend(index_entry["procedures"])
                        vba_index["globals"].extend(index_entry["globals"])
                    except Exception as e:
                        logger.warning("Skipping macro %s: %s", name, e)
        except Exception as e:
            logger.debug("Macro enumeration failed: %s", e)

    except Exception as e:
        logger.error("Access.Application extraction failed: %s", e)
    finally:
        if app is not None:
            _quit_access(app)

    try:
        index_path = output_dir / "index.json"
        index_path.write_text(
            json.dumps(vba_index, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception as e:
        logger.warning("Could not write VBA index.json: %s", e)

    return result
