"""Extract structured JSON report definitions from Access reports via COM.

Opens each report in Design View, iterates sections and controls, extracts all
properties with positions in twips, and recursively resolves subreports.
"""

import json
import logging
import subprocess
import time
from pathlib import Path

from .access_app import _get_access_app, _quit_access, _safe_filename

logger = logging.getLogger(__name__)

# Access report section indices
_SECTION_MAP = {
    0: "Detail",
    1: "ReportHeader",
    2: "ReportFooter",
    3: "PageHeader",
    4: "PageFooter",
    # 5,7,9... = GroupHeader0, GroupHeader1, GroupHeader2...
    # 6,8,10... = GroupFooter0, GroupFooter1, GroupFooter2...
}

# Access acControlType constants
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
    112: "SubReport",       # acSubform — used for both SubForm and SubReport
    114: "ObjectFrame",
    118: "PageBreak",
    122: "ToggleButton",
}

# SubReport/SubForm control type
_AC_SUBREPORT = 112

# acDesign = 1 (Design View for reports)
_AC_DESIGN = 1


def _is_com_broken(error: Exception) -> bool:
    """Detect COM errors that indicate the Access session is unusable.

    Covers RPC disconnects, server crashes, and Access-internal errors that
    leave the COM channel in a bad state (e.g. "The server threw an exception",
    "You can't carry out this action at the present time",
    "invalid reference to the property").
    """
    text = str(error).lower()
    markers = [
        "rpc server is unavailable",
        "-2147023174",
        "the server threw an exception",
        "-2147417851",
        "you can't carry out this action",
        "-2146825802",
        "invalid reference to the property",
        "-2146825833",
        "call was rejected by callee",
        "-2147418111",
    ]
    return any(m in text for m in markers)


def _kill_stale_access():
    """Kill any lingering MSACCESS.EXE processes to ensure a clean reconnection."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10,
        )
        logger.debug("Killed stale MSACCESS.EXE processes")
    except Exception as e:
        logger.debug("taskkill for MSACCESS.EXE: %s", e)


def _safe_getattr(obj, attr, default=None):
    """Safely read a COM property, returning default on failure."""
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def extract_reports(
    accdb_path: Path,
    output_dir: Path,
    report_names: list[str] | None = None,
    capture_screenshots: bool = False,
    vba_dir: Path | None = None,
    allow_direct_fallback: bool = False,
) -> dict:
    """
    Extract structured JSON for Access reports.

    Args:
        accdb_path: Path to the .accdb file.
        output_dir: Directory to write JSON files into.
        report_names: If provided, only extract these reports. Otherwise extract all.
        capture_screenshots: If True, capture designer/preview screenshots.
        vba_dir: Path to VBA reports directory (extract/vba/reports) for fallback
                 VBA code reading when COM Module access fails.

    Returns:
        {"reports": [...], "errors": [...], "count": int}
    """
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        logger.warning(
            "pywin32 not installed; cannot extract reports. "
            "Install with: pip install pywin32"
        )
        return {"reports": [], "errors": [], "count": 0}

    accdb_path = Path(accdb_path).resolve()
    if not accdb_path.exists():
        logger.error("Database not found: %s", accdb_path)
        return {"reports": [], "errors": [], "count": 0}

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    result = {"reports": [], "errors": [], "count": 0}

    app = None
    try:
        app = _get_access_app(
            str(accdb_path),
            allow_direct_fallback=allow_direct_fallback,
        )

        # Determine which reports to extract
        if report_names:
            names = report_names
        else:
            count = app.CurrentProject.AllReports.Count
            names = [app.CurrentProject.AllReports.Item(i).Name for i in range(count)]

        logger.info("Extracting %d reports...", len(names))

        consecutive_failures = 0
        _MAX_CONSECUTIVE_FAILURES = 5

        for name in names:
            max_attempts = 3
            extracted = False
            for attempt in range(1, max_attempts + 1):
                try:
                    visited = set()
                    report_data = _extract_single_report(
                        app, name, output_dir, visited, capture_screenshots, vba_dir
                    )
                    # Write JSON
                    json_path = output_dir / f"{_safe_filename(name)}.json"
                    json_path.write_text(
                        json.dumps(report_data, indent=2, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    result["reports"].append(report_data)
                    logger.info("  Extracted: %s", name)
                    extracted = True
                    consecutive_failures = 0
                    break
                except Exception as e:
                    if attempt < max_attempts and _is_com_broken(e):
                        logger.warning(
                            "  Access COM session broken while extracting %s (attempt %d); "
                            "killing stale processes and reconnecting...",
                            name, attempt,
                        )
                        try:
                            if app is not None:
                                _quit_access(app)
                        except Exception:
                            pass
                        _kill_stale_access()
                        time.sleep(2)
                        try:
                            app = _get_access_app(
                                str(accdb_path),
                                allow_direct_fallback=allow_direct_fallback,
                            )
                        except Exception as conn_err:
                            logger.warning(
                                "  Reconnection failed: %s; will retry", conn_err
                            )
                            _kill_stale_access()
                            time.sleep(3)
                            app = _get_access_app(
                                str(accdb_path),
                                allow_direct_fallback=allow_direct_fallback,
                            )
                        continue

                    logger.warning("  Failed: %s — %s", name, e)
                    result["errors"].append({"report": name, "error": str(e)})
                    break

            if not extracted:
                consecutive_failures += 1
                if consecutive_failures >= _MAX_CONSECUTIVE_FAILURES:
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
                    try:
                        app = _get_access_app(
                            str(accdb_path),
                            allow_direct_fallback=allow_direct_fallback,
                        )
                        consecutive_failures = 0
                    except Exception as conn_err:
                        logger.error(
                            "  Could not reconnect after consecutive failures: %s", conn_err
                        )
                        break

        result["count"] = len(result["reports"])

    except Exception as e:
        logger.error("Report extraction failed: %s", e)
        result["errors"].append({"report": "(connection)", "error": str(e)})
    finally:
        if app is not None:
            _quit_access(app)

    return result


def _extract_single_report(
    app, report_name: str, output_dir: Path, visited: set,
    capture_screenshots: bool, vba_dir: Path | None = None,
) -> dict:
    """Open a report in Design View and extract its full structure.

    Access locks subreports while the parent report is open, so we must:
    1. Open the parent, extract all metadata/sections/controls
    2. Close the parent
    3. Recursively extract each subreport
    4. Merge resolved subreport data back into the parent
    """
    if report_name in visited:
        return {"reportName": report_name, "_circular": True}
    visited.add(report_name)

    # Phase 1: Open report, extract everything except subreport recursion
    app.DoCmd.OpenReport(report_name, _AC_DESIGN)
    try:
        report = app.Reports(report_name)

        data = {"reportName": report_name}
        data.update(_get_report_metadata(report))

        groups = _get_grouping_info(report)
        if groups:
            data["groupLevels"] = groups

        data["sections"], pending_subreports = _extract_sections(report, groups)

        # Extract VBA code-behind from the report's Module (while still open)
        data["vbaCodeBehind"] = _extract_vba_code(report)

        # Fallback: read VBA from file if COM Module extraction returned nothing
        if data["vbaCodeBehind"] is None and vba_dir is not None:
            data["vbaCodeBehind"] = _read_vba_from_file(vba_dir, report_name)

        # Capture screenshots if requested (while report is still open)
        if capture_screenshots:
            try:
                from .screenshots import capture_report_screenshots

                screenshots_dir = output_dir / "screenshots"
                screenshots_dir.mkdir(parents=True, exist_ok=True)
                screenshot_paths = capture_report_screenshots(
                    app, report_name, screenshots_dir
                )
                if screenshot_paths:
                    data["screenshots"] = screenshot_paths
            except Exception as e:
                logger.debug("Screenshot capture failed for %s: %s", report_name, e)
    finally:
        try:
            app.DoCmd.Close(3, report_name, 0)  # acReport=3, acSaveNo=0
        except Exception:
            pass

    # Phase 2: Resolve subreports now that the parent is closed
    for ctrl_data, sub_report_name in pending_subreports:
        logger.debug("Resolving subreport: %s", sub_report_name)
        try:
            resolved = _extract_single_report(
                app, sub_report_name, output_dir, visited,
                capture_screenshots, vba_dir,
            )
            if resolved:
                ctrl_data["resolvedReport"] = resolved
        except Exception as e:
            logger.warning("Could not resolve subreport %s: %s", sub_report_name, e)
            ctrl_data["subreportError"] = str(e)

    return data


def _get_report_metadata(report) -> dict:
    """Extract report-level metadata."""
    meta = {}

    record_source = _safe_getattr(report, "RecordSource", "")
    if record_source:
        meta["dataSource"] = record_source

    width = _safe_getattr(report, "Width")
    if width is not None:
        meta["width"] = width

    caption = _safe_getattr(report, "Caption", "")
    if caption:
        meta["caption"] = caption

    # Page orientation: 1=Portrait, 2=Landscape
    printer = _safe_getattr(report, "Printer")
    if printer:
        orientation = _safe_getattr(printer, "Orientation")
        if orientation == 2:
            meta["orientation"] = "landscape"
        elif orientation == 1:
            meta["orientation"] = "portrait"

        # Margins (in twips)
        margins = {}
        for attr in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            val = _safe_getattr(printer, attr)
            if val is not None:
                key = attr[0].lower() + attr[1:]  # leftMargin, etc.
                # Simplify to just left, right, top, bottom
                key = attr.replace("Margin", "").lower()
                margins[key] = val
        if margins:
            meta["margins"] = margins

    return meta


def _get_grouping_info(report) -> list[dict]:
    """Extract GroupLevel collection."""
    groups = []
    for i in range(10):  # Access supports up to 10 group levels
        try:
            gl = report.GroupLevel(i)
            group = {
                "fieldExpression": _safe_getattr(gl, "ControlSource", "")
                or _safe_getattr(gl, "Expression", ""),
            }
            sort_order = _safe_getattr(gl, "SortOrder", 0)
            group["sortOrder"] = "descending" if sort_order else "ascending"
            group["groupHeader"] = bool(_safe_getattr(gl, "GroupHeader", False))
            group["groupFooter"] = bool(_safe_getattr(gl, "GroupFooter", False))
            groups.append(group)
        except Exception:
            break
    return groups


def _get_section_object(report, idx: int):
    """Try multiple COM access patterns to get a Section object."""
    # Pattern 1: Direct method call — report.Section(idx)
    try:
        return report.Section(idx)
    except Exception:
        pass

    # Pattern 2: Bracket/index syntax — report.Section[idx]
    try:
        return report.Section[idx]
    except Exception:
        pass

    # Pattern 3: Low-level COM invoke via _oleobj_
    try:
        oleobj = report._oleobj_
        import pythoncom
        return oleobj.Invoke(
            oleobj.GetIDsOfNames("Section")[0],
            0, pythoncom.DISPATCH_PROPERTYGET,
            True, idx,
        )
    except Exception:
        pass

    return None


def _extract_sections(report, groups: list[dict]) -> tuple[list[dict], list[tuple[dict, str]]]:
    """Iterate all report sections and extract controls.

    Controls are accessed via report.Controls (not section.Controls) because
    Access COM only populates the report-level collection. Each control has a
    .Section property indicating which section index it belongs to.

    Returns:
        (sections_list, pending_subreports) where pending_subreports is a list
        of (ctrl_data_dict, subreport_name) tuples to be resolved after the
        parent report is closed.
    """
    # Group all report-level controls by their section index
    controls_by_section: dict[int, list] = {}
    try:
        for i in range(report.Controls.Count):
            ctrl = report.Controls(i)
            sec_idx = _safe_getattr(ctrl, "Section", -1)
            controls_by_section.setdefault(sec_idx, []).append(ctrl)
    except Exception as e:
        logger.debug("Error iterating report controls: %s", e)

    # Determine which section indices exist: from controls + standard ones to probe
    known_indices = set(controls_by_section.keys()) | {0, 1, 2, 3, 4}
    # Add group header/footer indices based on group count
    for g in range(len(groups)):
        known_indices.add(5 + g * 2)  # GroupHeader
        known_indices.add(6 + g * 2)  # GroupFooter
    known_indices.discard(-1)

    sections = []
    pending_subreports = []

    for idx in sorted(known_indices):
        sec_data = _get_section_info(idx, groups)

        # Try to get section-level properties (Height, Visible, BackColor)
        section = _get_section_object(report, idx)
        if section is not None:
            sec_data["height"] = _safe_getattr(section, "Height", 0)
            visible = _safe_getattr(section, "Visible", True)
            if not visible:
                sec_data["visible"] = False
            back_color = _safe_getattr(section, "BackColor")
            if back_color is not None and back_color != 16777215:
                sec_data["backColor"] = back_color

        sec_controls = controls_by_section.get(idx, [])
        if not sec_controls and section is None:
            # Skip sections we can't access and have no controls
            continue

        extracted, sub_pending = _extract_controls(sec_controls)
        sec_data["controls"] = extracted
        pending_subreports.extend(sub_pending)

        sections.append(sec_data)

    return sections, pending_subreports


def _get_section_info(idx: int, groups: list[dict]) -> dict:
    """Determine section type and group info from index."""
    if idx in _SECTION_MAP:
        return {"type": _SECTION_MAP[idx]}

    # Group headers/footers
    if idx >= 5:
        group_idx = (idx - 5) // 2
        is_header = (idx - 5) % 2 == 0
        sec_type = "GroupHeader" if is_header else "GroupFooter"
        info = {"type": sec_type, "groupIndex": group_idx}
        if group_idx < len(groups):
            info["groupField"] = groups[group_idx].get("fieldExpression", "")
        return info

    return {"type": f"Unknown_{idx}"}


def _extract_controls(control_list: list) -> tuple[list[dict], list[tuple[dict, str]]]:
    """Extract properties from a list of COM control objects.

    Returns:
        (controls_list, pending_subreports) where pending_subreports contains
        (ctrl_data_dict, subreport_name) for deferred resolution.
    """
    controls = []
    pending_subreports = []
    for ctrl in control_list:
        try:
            ctrl_data = _extract_control_properties(ctrl)

            # Collect subreport references for deferred resolution
            ctrl_type_num = _safe_getattr(ctrl, "ControlType", -1)
            if ctrl_type_num == _AC_SUBREPORT:
                source_obj = _safe_getattr(ctrl, "SourceObject", "")
                if source_obj:
                    sub_name = source_obj
                    if sub_name.lower().startswith("report."):
                        sub_name = sub_name[7:]
                    pending_subreports.append((ctrl_data, sub_name))

            controls.append(ctrl_data)
        except Exception as e:
            logger.debug("Error extracting control: %s", e)
    return controls, pending_subreports


def _extract_control_properties(control) -> dict:
    """Extract properties from a single control."""
    ctrl_type_num = _safe_getattr(control, "ControlType", -1)
    ctrl_type_name = _CONTROL_TYPES.get(ctrl_type_num, f"Unknown_{ctrl_type_num}")

    props = {
        "type": ctrl_type_name,
        "name": _safe_getattr(control, "Name", ""),
    }

    # Position and size (twips)
    for attr in ("Left", "Top", "Width", "Height"):
        val = _safe_getattr(control, attr)
        if val is not None:
            props[attr[0].lower() + attr[1:]] = val

    # Content properties
    control_source = _safe_getattr(control, "ControlSource")
    if control_source:
        props["controlSource"] = control_source

    caption = _safe_getattr(control, "Caption")
    if caption:
        props["caption"] = caption

    # Font properties
    font_name = _safe_getattr(control, "FontName")
    if font_name:
        props["fontName"] = font_name

    font_size = _safe_getattr(control, "FontSize")
    if font_size is not None:
        props["fontSize"] = font_size

    font_bold = _safe_getattr(control, "FontBold")
    if font_bold:
        props["fontBold"] = True

    font_italic = _safe_getattr(control, "FontItalic")
    if font_italic:
        props["fontItalic"] = True

    font_underline = _safe_getattr(control, "FontUnderline")
    if font_underline:
        props["fontUnderline"] = True

    # Colors
    fore_color = _safe_getattr(control, "ForeColor")
    if fore_color is not None and fore_color != 0:  # not black
        props["foreColor"] = fore_color

    back_color = _safe_getattr(control, "BackColor")
    if back_color is not None and back_color != 16777215:  # not white
        props["backColor"] = back_color

    back_style = _safe_getattr(control, "BackStyle")
    if back_style is not None and back_style == 0:  # transparent
        props["backStyle"] = "transparent"

    # Borders
    border_style = _safe_getattr(control, "BorderStyle")
    if border_style is not None and border_style != 0:
        props["borderStyle"] = border_style

    border_width = _safe_getattr(control, "BorderWidth")
    if border_width is not None and border_width != 0:
        props["borderWidth"] = border_width

    border_color = _safe_getattr(control, "BorderColor")
    if border_color is not None and border_color != 0:
        props["borderColor"] = border_color

    # Format string
    format_str = _safe_getattr(control, "Format")
    if format_str:
        props["format"] = format_str

    # Text alignment: 0=General, 1=Left, 2=Center, 3=Right
    text_align = _safe_getattr(control, "TextAlign")
    if text_align is not None and text_align != 0:
        align_map = {1: "left", 2: "center", 3: "right"}
        props["textAlign"] = align_map.get(text_align, text_align)

    # CanGrow / CanShrink
    can_grow = _safe_getattr(control, "CanGrow")
    if can_grow:
        props["canGrow"] = True

    can_shrink = _safe_getattr(control, "CanShrink")
    if can_shrink:
        props["canShrink"] = True

    # SubReport-specific: SourceObject, LinkMasterFields, LinkChildFields
    if ctrl_type_num == _AC_SUBREPORT:
        source_obj = _safe_getattr(control, "SourceObject", "")
        if source_obj:
            props["sourceObject"] = source_obj

        link_master = _safe_getattr(control, "LinkMasterFields", "")
        if link_master:
            props["linkMasterFields"] = [f.strip() for f in link_master.split(";")]

        link_child = _safe_getattr(control, "LinkChildFields", "")
        if link_child:
            props["linkChildFields"] = [f.strip() for f in link_child.split(";")]

    # Line-specific properties
    if ctrl_type_num == 102:
        line_slant = _safe_getattr(control, "LineSlant")
        if line_slant:
            props["lineSlant"] = True

    # Visibility
    visible = _safe_getattr(control, "Visible", True)
    if not visible:
        props["visible"] = False

    return props


def _extract_vba_code(report) -> str | None:
    """Extract VBA code-behind from the report's COM Module object.

    Returns the VBA code string, or None if the report has no module or
    extraction fails.
    """
    try:
        has_module = _safe_getattr(report, "HasModule", False)
        if not has_module:
            return None

        module = report.Module
        count = module.CountOfLines
        if count == 0:
            return None

        code = module.Lines(1, count)
        return code if code and code.strip() else None
    except Exception as e:
        logger.debug("Could not read VBA module for report: %s", e)
        return None


def _read_vba_from_file(vba_dir: Path, report_name: str) -> str | None:
    """Read VBA code from a SaveAsText export file as fallback.

    Looks for <safe_filename(report_name)>.txt in vba_dir.
    Handles UTF-16LE encoding from Access SaveAsText.
    """
    file_path = vba_dir / f"{_safe_filename(report_name)}.txt"
    if not file_path.exists():
        return None

    try:
        # SaveAsText produces UTF-16LE files
        content = file_path.read_text(encoding="utf-16-le", errors="replace")
        return content if content and content.strip() else None
    except Exception as e:
        logger.debug("Could not read VBA file %s: %s", file_path, e)
        # Try UTF-8 as a secondary fallback
        try:
            content = file_path.read_text(encoding="utf-8", errors="replace")
            return content if content and content.strip() else None
        except Exception:
            return None
