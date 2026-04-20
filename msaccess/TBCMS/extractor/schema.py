"""Extract table schema from MS Access database via pyodbc."""

import json
import logging
from pathlib import Path

import pyodbc

logger = logging.getLogger(__name__)

# System tables to exclude
SYSTEM_TABLE_PREFIXES = ("MSys", "USys", "~", "Switchboard", ">")


def _is_user_table(table_name: str) -> bool:
    """Return True if table is a user table (not system)."""
    return not any(
        table_name.startswith(prefix) for prefix in SYSTEM_TABLE_PREFIXES
    )


def _get_primary_keys(cursor: pyodbc.Cursor, table_name: str) -> list[str]:
    """Get primary key column names for a table."""
    try:
        pk_columns = []
        for row in cursor.primaryKeys(table=table_name):
            pk_columns.append(row.column_name)
        return sorted(pk_columns)
    except pyodbc.Error:
        return []


def extract_schema(conn_str: str, output_path: Path) -> dict:
    """
    Extract table schema to JSON.
    Returns the schema dict for use by other modules.
    """
    schema = {"tables": []}
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except pyodbc.Error as e:
        logger.error("Failed to connect: %s", e)
        raise

    try:
        tables = []
        for row in cursor.tables(tableType="TABLE"):
            try:
                name = row.table_name
                if _is_user_table(name):
                    tables.append(str(name) if name else "")
            except (UnicodeDecodeError, UnicodeEncodeError):
                continue
        tables = [t for t in tables if t]

        for table_name in tables:
            try:
                columns = []
                for col in cursor.columns(table=table_name):
                    columns.append({
                        "name": str(col.column_name) if col.column_name else "",
                        "type": str(col.type_name) if col.type_name else "",
                        "size": col.column_size,
                    })
                primary_key = _get_primary_keys(cursor, table_name)
                schema["tables"].append({
                    "name": str(table_name),
                    "columns": columns,
                    "primary_key": primary_key,
                })
            except (pyodbc.Error, UnicodeDecodeError, UnicodeEncodeError) as e:
                logger.warning("Skipping table %s: %s", table_name, e)

    finally:
        conn.close()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(schema, f, indent=2)

    return schema
