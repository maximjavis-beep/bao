"""报关记录 CRUD 操作"""

import json
import uuid
from datetime import date, datetime
from typing import Optional

from .connection import Database
from .schema import ALL_DDL


class DeclarationRepository:
    """出口报关单数据访问"""

    def __init__(self, db: Database):
        self.db = db
        self._init_db()

    def _init_db(self):
        for ddl in ALL_DDL:
            self.db.execute(ddl)
        self.db.commit()

    def save(self, declaration) -> str:
        """保存报关单（已废弃 — FBA 模式下不再使用）"""
        return ""

    def list_all(self, limit: int = 50, offset: int = 0) -> list[dict]:
        """列出报关单"""
        rows = self.db.execute(
            """
            SELECT * FROM declarations
            ORDER BY created_at DESC LIMIT ? OFFSET ?
            """,
            (limit, offset),
        ).fetchall()
        return [dict(r) for r in rows]

    def search(self, keyword: str) -> list[dict]:
        """模糊搜索"""
        pattern = f"%{keyword}%"
        rows = self.db.execute(
            """
            SELECT * FROM declarations
            WHERE declaration_id LIKE ?
               OR domestic_shipper LIKE ?
               OR overseas_consignee LIKE ?
               OR destination_country LIKE ?
               OR contract_no LIKE ?
               OR invoice_no LIKE ?
            ORDER BY created_at DESC
            LIMIT 50
            """,
            (pattern, pattern, pattern, pattern, pattern, pattern),
        ).fetchall()
        return [dict(r) for r in rows]

    def get_items(self, declaration_id: str) -> list[dict]:
        """获取某报关单的商品明细"""
        rows = self.db.execute(
            "SELECT * FROM declaration_items WHERE declaration_fk = ? ORDER BY seq",
            (declaration_id,),
        ).fetchall()
        return [dict(r) for r in rows]

    def stats_by_month(self) -> list[dict]:
        """按月统计"""
        rows = self.db.execute(
            """
            SELECT
                substr(declaration_date, 1, 7) AS month,
                COUNT(*) AS count,
                SUM(total_amount) AS total_amount,
                SUM(fob_amount) AS total_fob,
                SUM(estimated_tax_rebate) AS total_rebate
            FROM declarations
            GROUP BY month
            ORDER BY month DESC
            """
        ).fetchall()
        return [dict(r) for r in rows]
