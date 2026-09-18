from __future__ import annotations

import sqlite3
from contextlib import closing
from pathlib import Path

DEFAULT_DB_PATH = Path("data/office_docs_mcp.db")


def initialize_database(path: str | Path = DEFAULT_DB_PATH) -> Path:
    db_path = Path(path)
    db_path.parent.mkdir(parents=True, exist_ok=True)

    with closing(sqlite3.connect(db_path)) as connection:
        with connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS app_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )
            connection.execute(
                "INSERT OR IGNORE INTO app_metadata(key, value) VALUES (?, ?)",
                ("created_by", "wpygen"),
            )

    return db_path
