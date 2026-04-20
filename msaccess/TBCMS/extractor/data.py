"""Export table data from MS Access to CSV via pyodbc."""

import csv
import logging
from pathlib import Path

import pyodbc

from .schema import SYSTEM_TABLE_PREFIXES

logger = logging.getLogger(__name__)

# Characters invalid in Windows filenames
INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def _safe_table_filename(name: str) -> str:
    """Convert table name to safe filename for CSV."""
    return "".join(
        c if c not in INVALID_FILENAME_CHARS else "_" for c in name
    )


def _is_user_table(table_name: str) -> bool:
    """Return True if table is a user table (not system)."""
    return not any(
        table_name.startswith(prefix) for prefix in SYSTEM_TABLE_PREFIXES
    )


def export_table_data(
    conn_str: str, output_dir: Path, table_names: list[str] | None = None
) -> list[str]:
    """
    Export each user table to CSV.
    If table_names is None, export all user tables.
    Returns list of exported table names.
    """
    exported = []
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except pyodbc.Error as e:
        logger.error("Failed to connect: %s", e)
        raise

    try:
        if table_names is None:
            tables = [
                row.table_name
                for row in cursor.tables(tableType="TABLE")
                if _is_user_table(row.table_name)
            ]
        else:
            tables = table_names

        for table_name in tables:
            try:
                cursor.execute(f"SELECT * FROM [{table_name}]")
                rows = cursor.fetchall()
                columns = [desc[0] for desc in cursor.description]

                safe_name = _safe_table_filename(table_name)
                output_path = output_dir / f"{safe_name}.csv"
                output_path.parent.mkdir(parents=True, exist_ok=True)

                with open(
                    output_path, "w", newline="", encoding="utf-8", errors="replace"
                ) as f:
                    writer = csv.writer(f)
                    writer.writerow(columns)
                    # Encode row values safely (handle bytes/encoding issues)
                    for row in rows:
                        safe_row = []
                        for v in row:
                            if v is None:
                                safe_row.append("")
                            elif isinstance(v, bytes):
                                safe_row.append(
                                    v.decode("utf-8", errors="replace")
                                )
                            else:
                                safe_row.append(str(v))
                        writer.writerow(safe_row)

                exported.append(table_name)
            except (pyodbc.Error, UnicodeDecodeError, UnicodeEncodeError) as e:
                logger.warning("Skipping table %s: %s", table_name, e)

    finally:
        conn.close()

    return exported
