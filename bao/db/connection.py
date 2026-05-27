"""SQLite 连接管理"""

import sqlite3
from pathlib import Path
from typing import Optional


class Database:
    """SQLite 数据库封装"""

    def __init__(self, db_path: str = ""):
        if db_path:
            self.db_path = Path(db_path)
        else:
            # 默认路径: customs/bao/bao/data/bao.db
            self.db_path = Path(__file__).parent.parent.parent / "data" / "bao.db"
        self._conn: Optional[sqlite3.Connection] = None

    @property
    def conn(self) -> sqlite3.Connection:
        if self._conn is None:
            self._conn = sqlite3.connect(str(self.db_path))
            self._conn.row_factory = sqlite3.Row
            self._conn.execute("PRAGMA journal_mode=WAL")
            self._conn.execute("PRAGMA foreign_keys=ON")
        return self._conn

    def close(self):
        if self._conn:
            self._conn.close()
            self._conn = None

    def execute(self, sql: str, params=()):
        return self.conn.execute(sql, params)

    def executemany(self, sql: str, params):
        return self.conn.executemany(sql, params)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()
