"""Shared Access.Application COM connection utilities.

Provides two strategies to connect to Access:
1. Subprocess launch + GetActiveObject (default, uses /nostartup)
2. Direct COM Dispatch fallback (opt-in)

Also provides a context manager for safe COM lifecycle management and a shared
safe-filename helper.
"""

import logging
import subprocess
import time
import winreg
from contextlib import contextmanager
from pathlib import Path

logger = logging.getLogger(__name__)

# Track subprocess handles separately (can't set attrs on COM objects in newer pywin32)
_subprocess_handles: dict[int, subprocess.Popen] = {}

# Retry settings for GetActiveObject fallback
_LAUNCH_WAIT_SECONDS = 15
_POLL_INTERVAL_SECONDS = 3
_POLL_MAX_ATTEMPTS = 15
_LOCK_MARKERS = (
    "prevents it from being opened or locked",
    "placed in a state by",
    "(-3810)",
)


def _is_database_locked(accdb_path: str) -> tuple[bool, str | None]:
    """Best-effort lock detection using ODBC before COM open attempts.

    Some locked database states can cause Access COM OpenCurrentDatabase to hang
    with no exception. This precheck lets us fail fast with a clear message.
    """
    try:
        import pyodbc
    except Exception:
        # If pyodbc is unavailable, keep existing behavior.
        return False, None

    conn_str = (
        f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};"
        f"DBQ={Path(accdb_path).resolve()};"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
        conn.close()
        return False, None
    except Exception as e:
        text = str(e).lower()
        if any(marker in text for marker in _LOCK_MARKERS):
            return True, str(e)
        return False, None


def _find_msaccess_exe() -> str | None:
    """Locate MSACCESS.EXE from the COM registry."""
    clsid_paths = [
        (winreg.HKEY_CLASSES_ROOT, r"CLSID\{73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}\LocalServer32"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\CLSID\{73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}\LocalServer32"),
    ]
    for root, subkey in clsid_paths:
        try:
            key = winreg.OpenKey(root, subkey)
            val = winreg.QueryValueEx(key, "")[0]
            winreg.CloseKey(key)
            exe_path = val.strip().strip('"')
            if Path(exe_path).exists():
                return exe_path
        except (FileNotFoundError, OSError):
            continue
    return None


def _connect_dispatch(accdb_path: str):
    """Try direct COM Dispatch to Access.Application."""
    import pythoncom
    import win32com.client

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.OpenCurrentDatabase(accdb_path, False)
    _ = app.CurrentProject.AllForms.Count
    return app


def _kill_all_access():
    """Kill any lingering MSACCESS.EXE processes to ensure a clean slate."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10,
        )
        logger.debug("Killed stale MSACCESS.EXE processes")
    except Exception as e:
        logger.debug("taskkill for MSACCESS.EXE: %s", e)


def _connect_subprocess(accdb_path: str):
    """Launch MSACCESS.EXE as a subprocess and connect via GetActiveObject.

    Kills any stale Access processes first to avoid connecting to a broken
    instance from a previous failed session.
    """
    import pythoncom
    import win32com.client

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    exe = _find_msaccess_exe()
    if not exe:
        raise RuntimeError("Cannot find MSACCESS.EXE in registry")

    _kill_all_access()
    time.sleep(2)

    logger.info("Launching Access subprocess with startup bypass: %s", exe)
    # /nostartup suppresses startup options (AutoExec/startup form behavior).
    proc = subprocess.Popen([exe, accdb_path, "/nostartup"])
    logger.info("Waiting %ds for Access to load...", _LAUNCH_WAIT_SECONDS)
    time.sleep(_LAUNCH_WAIT_SECONDS)

    app = None
    for attempt in range(1, _POLL_MAX_ATTEMPTS + 1):
        try:
            app = win32com.client.GetActiveObject("Access.Application")
            _ = app.CurrentProject.AllForms.Count
            logger.info("Connected to Access on attempt %d", attempt)
            break
        except Exception:
            logger.debug("GetActiveObject attempt %d failed, retrying...", attempt)
            time.sleep(_POLL_INTERVAL_SECONDS)

    if app is None:
        proc.terminate()
        raise RuntimeError(
            f"Failed to connect to Access.Application after {_POLL_MAX_ATTEMPTS} attempts"
        )

    _subprocess_handles[id(app)] = proc
    return app


def _get_access_app(accdb_path: str, allow_direct_fallback: bool = False):
    """Get a connected Access.Application instance using the best available method.

    If direct COM Dispatch fails (e.g. database already open in a stale process),
    kills all Access processes before trying the subprocess fallback.
    """
    locked, lock_error = _is_database_locked(accdb_path)
    if locked:
        raise RuntimeError(
            "Database appears locked by another Access/ODBC session; "
            f"cannot open via COM safely. Details: {lock_error}"
        )

    # Prefer subprocess strategy for stability and startup suppression.
    try:
        return _connect_subprocess(accdb_path)
    except Exception as subproc_error:
        if not allow_direct_fallback:
            raise RuntimeError(
                "Access subprocess connection failed while enforcing startup suppression "
                "(/nostartup). Direct COM fallback is disabled."
            ) from subproc_error
        logger.info(
            "Subprocess Access connection failed (%s), trying direct COM Dispatch...",
            subproc_error,
        )
        return _connect_dispatch(accdb_path)


def _quit_access(app):
    """Cleanly quit Access and terminate any subprocess."""
    try:
        app.Quit()
    except Exception:
        pass

    proc = _subprocess_handles.pop(id(app), None)
    if proc is not None:
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            logger.warning("Access process did not exit; terminating forcefully")
            proc.terminate()


@contextmanager
def access_app_context(accdb_path: str | Path):
    """Context manager for safe Access.Application COM lifecycle.

    Usage:
        with access_app_context(accdb_path) as app:
            # use app...
    """
    app = _get_access_app(str(accdb_path))
    try:
        yield app
    finally:
        _quit_access(app)


def _safe_filename(name: str) -> str:
    """Convert object name to safe filename."""
    return "".join(
        c if c.isalnum() or c in "._-" else "_" for c in name
    )
