"""Optional screenshot capture for Access reports.

Captures designer (Design View) and preview (Print Preview) screenshots of
Access reports via the pywin32 window capture API. Falls back to BMP format
if Pillow is not installed.
"""

import logging
import time
from pathlib import Path

from .access_app import _safe_filename

logger = logging.getLogger(__name__)

# acDesign=1, acPrintPreview=5, acReport=3, acSaveNo=0
_AC_DESIGN = 1
_AC_PRINT_PREVIEW = 5
_AC_REPORT = 3
_AC_SAVE_NO = 0


def _capture_window(hwnd, output_path: Path) -> Path:
    """Capture a window to an image file.

    Saves as BMP natively; converts to PNG if Pillow is available.
    Returns the actual path written (may differ in extension).
    """
    import win32con
    import win32gui
    import win32ui

    # Get window dimensions
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    if width <= 0 or height <= 0:
        raise ValueError(f"Invalid window dimensions: {width}x{height}")

    # Create device contexts and bitmap
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)

    # BitBlt the window content
    save_dc.BitBlt(
        (0, 0), (width, height),
        mfc_dc, (0, 0),
        win32con.SRCCOPY,
    )

    # Save as BMP first
    bmp_path = output_path.with_suffix(".bmp")
    bitmap.SaveBitmapFile(save_dc, str(bmp_path))

    # Clean up GDI resources
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    win32gui.DeleteObject(bitmap.GetHandle())

    # Try to convert to PNG with Pillow
    try:
        from PIL import Image

        img = Image.open(str(bmp_path))
        png_path = output_path.with_suffix(".png")
        img.save(str(png_path))
        img.close()
        bmp_path.unlink(missing_ok=True)
        return png_path
    except ImportError:
        logger.debug("Pillow not installed; screenshot saved as BMP")
        return bmp_path


def capture_report_screenshots(
    app, report_name: str, screenshots_dir: Path
) -> dict[str, str]:
    """Capture designer and preview screenshots for a report.

    The report should already be open in Design View when this is called.

    Args:
        app: Access.Application COM object.
        report_name: Name of the report.
        screenshots_dir: Directory to save screenshots.

    Returns:
        Dict with "designer" and/or "preview" keys mapped to file paths.
    """
    import win32gui

    screenshots_dir.mkdir(parents=True, exist_ok=True)
    safe_name = _safe_filename(report_name)
    result = {}

    hwnd = app.hWndAccessApp()

    # Bring window to foreground for capture
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)
    except Exception:
        pass

    # Designer screenshot (report should already be open in design view)
    try:
        designer_path = screenshots_dir / f"{safe_name}_designer"
        actual_path = _capture_window(hwnd, designer_path)
        result["designer"] = str(actual_path)
        logger.debug("  Designer screenshot: %s", actual_path.name)
    except Exception as e:
        logger.debug("  Designer screenshot failed: %s", e)

    # Preview screenshot: close design view, reopen in Print Preview
    try:
        app.DoCmd.Close(_AC_REPORT, report_name, _AC_SAVE_NO)
        app.DoCmd.OpenReport(report_name, _AC_PRINT_PREVIEW)
        time.sleep(1)  # Wait for preview to render

        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception:
            pass

        preview_path = screenshots_dir / f"{safe_name}_preview"
        actual_path = _capture_window(hwnd, preview_path)
        result["preview"] = str(actual_path)
        logger.debug("  Preview screenshot: %s", actual_path.name)

        # Close preview and reopen in design view for caller
        app.DoCmd.Close(_AC_REPORT, report_name, _AC_SAVE_NO)
        app.DoCmd.OpenReport(report_name, _AC_DESIGN)
    except Exception as e:
        logger.debug("  Preview screenshot failed: %s", e)
        # Try to reopen in design view
        try:
            app.DoCmd.OpenReport(report_name, _AC_DESIGN)
        except Exception:
            pass

    return result
