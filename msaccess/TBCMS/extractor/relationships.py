"""Extract table relationships from MS Access database via pyodbc."""

import json
import logging
from pathlib import Path

import pyodbc

logger = logging.getLogger(__name__)


def extract_relationships(conn_str: str, output_path: Path) -> list[dict]:
    """
    Extract foreign key relationships.
    Tries MSysRelationships first; falls back to ODBC foreignKeys if needed.
    Returns the relationships list.
    """
    relationships = []
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except pyodbc.Error as e:
        logger.error("Failed to connect: %s", e)
        raise

    try:
        # Try ODBC foreignKeys metadata
        tables = [
            row.table_name
            for row in cursor.tables(tableType="TABLE")
            if not row.table_name.startswith("MSys")
            and not row.table_name.startswith("~")
        ]

        for table_name in tables:
            try:
                for fk in cursor.foreignKeys(table=table_name):
                    # pyodbc uses FKTABLE_NAME, FKCOLUMN_NAME, etc. (ODBC standard)
                    rel = {
                        "table": getattr(fk, "fk_table_name", None)
                        or getattr(fk, "FKTABLE_NAME", None),
                        "columns": getattr(fk, "fk_column_name", None)
                        or getattr(fk, "FKCOLUMN_NAME", None),
                        "referenced_table": getattr(fk, "pk_table_name", None)
                        or getattr(fk, "PKTABLE_NAME", None),
                        "referenced_columns": getattr(fk, "pk_column_name", None)
                        or getattr(fk, "PKCOLUMN_NAME", None),
                    }
                    if rel["table"] is None or rel["referenced_table"] is None:
                        continue
                    if isinstance(rel["columns"], str):
                        rel["columns"] = [rel["columns"]]
                    if isinstance(rel["referenced_columns"], str):
                        rel["referenced_columns"] = [rel["referenced_columns"]]
                    relationships.append(rel)
            except pyodbc.Error:
                pass

        # Try MSysRelationships for additional metadata (may require permissions)
        try:
            cursor.execute(
                """
                SELECT szObject, szReferencedObject, szColumn, szReferencedColumn
                FROM MSysRelationships
                WHERE (szObject NOT LIKE 'MSys%' AND szObject NOT LIKE '~%')
                """
            )
            for row in cursor.fetchall():
                rel = {
                    "table": row.szObject,
                    "referenced_table": row.szReferencedObject,
                    "columns": [row.szColumn] if row.szColumn else [],
                    "referenced_columns": (
                        [row.szReferencedColumn] if row.szReferencedColumn else []
                    ),
                }
                if rel not in relationships:
                    relationships.append(rel)
        except pyodbc.Error:
            pass

    finally:
        conn.close()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(relationships, f, indent=2)

    return relationships
