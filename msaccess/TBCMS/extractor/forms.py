"""Extract structured JSON form definitions from Access forms via COM."""

import json
import logging
import multiprocessing as mp
import subprocess
import time
from pathlib import Path

from .access_app import _get_access_app, _quit_access, _safe_filename

logger = logging.getLogger(__name__)

_SECTION_MAP = {
    0: "Detail",
    1: "FormHeader",
    2: "FormFooter",
    3: "PageHeader",
    4: "PageFooter",
}

_CONTROL_TYPES = {
    100: "Label",
    101: "Rectangle",
    102: "Line",
    103: "Image",
    104: "CommandButton",
    105: "OptionButton",
    106: "CheckBox",
    107: "OptionGroup",
    108: "BoundObjectFrame",
    109: "TextBox",
    110: "ListBox",
    111: "ComboBox",
    112: "SubForm",
    114: "ObjectFrame",
    118: "PageBreak",
    122: "ToggleButton",
}

_AC_FORM = 2
_AC_DESIGN = 1
_AC_SUBFORM = 112
_FORM_TIMEOUT_SECONDS = 60


def _safe_getattr(obj, attr, default=None):
    """Safely read a COM property, returning default on failure."""
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _is_com_broken(error: Exception) -> bool:
    """Detect COM errors likely requiring Access reconnect."""
    text = str(error).lower()
    markers = [
        "rpc server is unavailable",
        "the server threw an exception",
        "call was rejected by callee",
        "you can't carry out this action",
    ]
    return any(m in text for m in markers)


def _kill_stale_access():
    """Kill any lingering MSACCESS.EXE processes to ensure a clean reconnection."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True,
            timeout=10,
        )
        logger.debug("Killed stale MSACCESS.EXE processes")
    except Exception as e:
        logger.debug("taskkill for MSACCESS.EXE: %s", e)


def extract_forms(
    accdb_path: Path,
    output_dir: Path,
    form_names: list[str] | None = None,
    vba_dir: Path | None = None,
    resolve_subforms: bool = False,
    per_form_timeout_seconds: int = _FORM_TIMEOUT_SECONDS,
    allow_direct_fallback: bool = False,
) -> dict:
    """Extract structured JSON for Access forms.

    Uses a single Access session for the full form list, reconnecting only
    when COM becomes unusable. This mirrors the report extraction strategy.
    """
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        logger.warning(
            "pywin32 not installed; cannot extract forms. "
            "Install with: pip install pywin32"
        )
        return {"forms": [], "errors": [], "count": 0}

    accdb_path = Path(accdb_path).resolve()
    if not accdb_path.exists():
        logger.error("Database not found: %s", accdb_path)
        return {"forms": [], "errors": [], "count": 0}

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    result = {"forms": [], "errors": [], "count": 0}
    app = None
    try:
        app = _get_access_app(
            str(accdb_path),
            allow_direct_fallback=allow_direct_fallback,
        )

        if form_names:
            names = form_names
        else:
            count = app.CurrentProject.AllForms.Count
            names = [app.CurrentProject.AllForms.Item(i).Name for i in range(count)]

        logger.info("Extracting %d forms...", len(names))
        if per_form_timeout_seconds != _FORM_TIMEOUT_SECONDS:
            logger.info(
                "Per-form timeout (%ds) is ignored in single-session form extraction mode.",
                per_form_timeout_seconds,
            )

        consecutive_failures = 0
        max_consecutive_failures = 5

        total = len(names)
        for idx, name in enumerate(names, start=1):
            logger.info("  Form %d/%d: %s", idx, total, name)
            max_attempts = 3
            extracted = False
            for attempt in range(1, max_attempts + 1):
                try:
                    visited = set()
                    form_data = _extract_single_form(
                        app=app,
                        form_name=name,
                        visited=visited,
                        vba_dir=vba_dir,
                        resolve_subforms=resolve_subforms,
                    )
                    json_path = output_dir / f"{_safe_filename(name)}.json"
                    json_path.write_text(
                        json.dumps(form_data, indent=2, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    result["forms"].append(form_data)
                    extracted = True
                    consecutive_failures = 0
                    break
                except Exception as e:
                    if attempt < max_attempts and _is_com_broken(e):
                        logger.warning(
                            "  Access COM session broken while extracting %s (attempt %d); "
                            "killing stale processes and reconnecting...",
                            name,
                            attempt,
                        )
                        try:
                            if app is not None:
                                _quit_access(app)
                        except Exception:
                            pass
                        _kill_stale_access()
                        time.sleep(2)
                        app = _get_access_app(
                            str(accdb_path),
                            allow_direct_fallback=allow_direct_fallback,
                        )
                        continue

                    logger.warning("  Failed: %s — %s", name, e)
                    result["errors"].append({"form": name, "error": str(e)})
                    break

            if not extracted:
                consecutive_failures += 1
                if consecutive_failures >= max_consecutive_failures:
                    logger.warning(
                        "  %d consecutive failures — reconnecting before continuing...",
                        consecutive_failures,
                    )
                    try:
                        if app is not None:
                            _quit_access(app)
                    except Exception:
                        pass
                    _kill_stale_access()
                    time.sleep(2)
                    app = _get_access_app(
                        str(accdb_path),
                        allow_direct_fallback=allow_direct_fallback,
                    )
                    consecutive_failures = 0

        result["count"] = len(result["forms"])
        return result
    except Exception as e:
        logger.error("Form extraction failed: %s", e)
        result["errors"].append({"form": "(connection)", "error": str(e)})
        return result
    finally:
        if app is not None:
            _quit_access(app)


def _extract_form_worker(
    accdb_path_str: str,
    form_name: str,
    vba_dir_str: str | None,
    resolve_subforms: bool,
    queue,
):
    """Child worker: extract a single form in isolated process."""
    app = None
    try:
        vba_dir = Path(vba_dir_str) if vba_dir_str else None
        app = _get_access_app(accdb_path_str)
        visited = set()
        data = _extract_single_form(app, form_name, visited, vba_dir, resolve_subforms)
        queue.put({"ok": True, "data": data})
    except Exception as e:
        queue.put({"ok": False, "error": str(e)})
    finally:
        if app is not None:
            _quit_access(app)


def _extract_single_form_with_timeout(
    accdb_path: Path,
    form_name: str,
    vba_dir: Path | None,
    resolve_subforms: bool,
    timeout_seconds: int,
) -> tuple[dict | None, str | None, bool]:
    """Extract one form with hard timeout via subprocess isolation."""
    ctx = mp.get_context("spawn")
    queue = ctx.Queue()
    proc = ctx.Process(
        target=_extract_form_worker,
        args=(
            str(accdb_path),
            form_name,
            str(vba_dir) if vba_dir is not None else None,
            resolve_subforms,
            queue,
        ),
    )
    proc.start()
    proc.join(timeout_seconds)

    if proc.is_alive():
        proc.terminate()
        proc.join(5)
        return None, None, True

    payload = None
    try:
        if not queue.empty():
            payload = queue.get_nowait()
    except Exception:
        payload = None
    finally:
        queue.close()
        queue.join_thread()

    if payload and payload.get("ok"):
        return payload.get("data"), None, False
    if payload and not payload.get("ok"):
        return None, payload.get("error", "unknown error"), False

    if proc.exitcode not in (0, None):
        return None, f"worker exited with code {proc.exitcode}", False
    return None, "no result from worker", False


def _extract_single_form(
    app, form_name: str, visited: set, vba_dir: Path | None = None, resolve_subforms: bool = False
) -> dict:
    """Open form in design view, extract metadata/sections/controls, resolve subforms."""
    if form_name in visited:
        return {"formName": form_name, "_circular": True}
    visited.add(form_name)

    app.DoCmd.OpenForm(form_name, _AC_DESIGN)
    try:
        form = app.Forms(form_name)

        data = {"formName": form_name}
        data.update(_get_form_metadata(form))
        data["sections"], pending_subforms = _extract_sections(form)
        data["formEvents"] = _extract_form_events(form)
        data["vbaCodeBehind"] = _extract_vba_code(form)
        if data["vbaCodeBehind"] is None and vba_dir is not None:
            data["vbaCodeBehind"] = _read_vba_from_file(vba_dir, form_name)
    finally:
        try:
            app.DoCmd.Close(_AC_FORM, form_name, 0)  # acSaveNo
        except Exception:
            pass

    if resolve_subforms:
        for ctrl_data, subform_name in pending_subforms:
            try:
                resolved = _extract_single_form(app, subform_name, visited, vba_dir, resolve_subforms)
                if resolved:
                    ctrl_data["resolvedSubform"] = resolved
            except Exception as e:
                logger.warning("Could not resolve subform %s: %s", subform_name, e)
                ctrl_data["subformError"] = str(e)
    elif pending_subforms:
        data["subformReferences"] = sorted({name for _, name in pending_subforms if name})

    return data


def _get_form_metadata(form) -> dict:
    """Extract form-level metadata."""
    meta = {}
    for attr, key in (
        ("Caption", "caption"),
        ("RecordSource", "recordSource"),
        ("DefaultView", "defaultView"),
        ("AllowEdits", "allowEdits"),
        ("AllowAdditions", "allowAdditions"),
        ("AllowDeletions", "allowDeletions"),
        ("DataEntry", "dataEntry"),
        ("NavigationButtons", "navigationButtons"),
        ("Width", "width"),
    ):
        value = _safe_getattr(form, attr)
        if value is not None and value != "":
            meta[key] = value
    return meta


def _extract_form_events(form) -> dict:
    """Extract common form-level event bindings."""
    events = {}
    for attr, key in (
        ("OnOpen", "onOpen"),
        ("OnLoad", "onLoad"),
        ("OnCurrent", "onCurrent"),
        ("OnBeforeUpdate", "onBeforeUpdate"),
        ("OnAfterUpdate", "onAfterUpdate"),
        ("OnClose", "onClose"),
    ):
        value = _safe_getattr(form, attr)
        if value:
            events[key] = value
    return events


def _extract_sections(form) -> tuple[list[dict], list[tuple[dict, str]]]:
    """Extract form sections and controls."""
    controls_by_section: dict[int, list] = {}
    try:
        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            sec_idx = _safe_getattr(ctrl, "Section", -1)
            controls_by_section.setdefault(sec_idx, []).append(ctrl)
    except Exception as e:
        logger.debug("Error iterating form controls: %s", e)

    known_indices = set(controls_by_section.keys()) | {0, 1, 2, 3, 4}
    known_indices.discard(-1)

    sections = []
    pending_subforms = []
    for idx in sorted(known_indices):
        sec_controls = controls_by_section.get(idx, [])
        if not sec_controls:
            continue
        sec_data = {"type": _SECTION_MAP.get(idx, f"Unknown_{idx}")}
        extracted, sub_pending = _extract_controls(sec_controls)
        sec_data["controls"] = extracted
        pending_subforms.extend(sub_pending)
        sections.append(sec_data)

    return sections, pending_subforms


def _extract_controls(control_list: list) -> tuple[list[dict], list[tuple[dict, str]]]:
    """Extract properties for controls in one section."""
    controls = []
    pending_subforms = []
    for ctrl in control_list:
        try:
            ctrl_data = _extract_control_properties(ctrl)
            ctrl_type_num = _safe_getattr(ctrl, "ControlType", -1)
            if ctrl_type_num == _AC_SUBFORM:
                source_obj = _safe_getattr(ctrl, "SourceObject", "")
                if source_obj:
                    sub_name = source_obj
                    if sub_name.lower().startswith("form."):
                        sub_name = sub_name[5:]
                    pending_subforms.append((ctrl_data, sub_name))
            controls.append(ctrl_data)
        except Exception as e:
            logger.debug("Error extracting form control: %s", e)
    return controls, pending_subforms


def _extract_control_properties(control) -> dict:
    """Extract a normalized subset of control properties."""
    ctrl_type_num = _safe_getattr(control, "ControlType", -1)
    ctrl_type_name = _CONTROL_TYPES.get(ctrl_type_num, f"Unknown_{ctrl_type_num}")
    props = {
        "type": ctrl_type_name,
        "name": _safe_getattr(control, "Name", ""),
    }

    for attr in ("Left", "Top", "Width", "Height"):
        value = _safe_getattr(control, attr)
        if value is not None:
            props[attr[0].lower() + attr[1:]] = value

    for attr, key in (
        ("ControlSource", "controlSource"),
        ("RowSource", "rowSource"),
        ("Caption", "caption"),
        ("Format", "format"),
    ):
        value = _safe_getattr(control, attr)
        if value:
            props[key] = value

    tab_index = _safe_getattr(control, "TabIndex")
    if tab_index is not None:
        props["tabIndex"] = tab_index

    for attr, key in (("Visible", "visible"), ("Enabled", "enabled"), ("Locked", "locked")):
        value = _safe_getattr(control, attr)
        if value is not None:
            props[key] = bool(value)

    events = {}
    for attr, key in (
        ("OnClick", "onClick"),
        ("OnDblClick", "onDblClick"),
        ("OnChange", "onChange"),
        ("OnBeforeUpdate", "onBeforeUpdate"),
        ("OnAfterUpdate", "onAfterUpdate"),
    ):
        value = _safe_getattr(control, attr)
        if value:
            events[key] = value
    if events:
        props["events"] = events

    if ctrl_type_num == _AC_SUBFORM:
        source_obj = _safe_getattr(control, "SourceObject", "")
        if source_obj:
            props["sourceObject"] = source_obj
        link_master = _safe_getattr(control, "LinkMasterFields", "")
        if link_master:
            props["linkMasterFields"] = [f.strip() for f in link_master.split(";")]
        link_child = _safe_getattr(control, "LinkChildFields", "")
        if link_child:
            props["linkChildFields"] = [f.strip() for f in link_child.split(";")]

    return props


def _extract_vba_code(form) -> str | None:
    """Extract form code-behind from COM module."""
    try:
        has_module = _safe_getattr(form, "HasModule", False)
        if not has_module:
            return None
        module = form.Module
        count = module.CountOfLines
        if count == 0:
            return None
        code = module.Lines(1, count)
        return code if code and code.strip() else None
    except Exception as e:
        logger.debug("Could not read VBA module for form: %s", e)
        return None


def _read_vba_from_file(vba_dir: Path, form_name: str) -> str | None:
    """Read form VBA from SaveAsText export file as fallback."""
    file_path = vba_dir / f"{_safe_filename(form_name)}.txt"
    if not file_path.exists():
        return None
    for encoding in ("utf-16-le", "utf-8"):
        try:
            content = file_path.read_text(encoding=encoding, errors="replace")
            return content if content and content.strip() else None
        except Exception:
            continue
    return None
