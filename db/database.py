"""
SQLite 資料庫管理
支援交易記錄、預算設定、匯入歷程、系統設定
"""

import json
import logging
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

DB_PATH = Path(__file__).parent.parent / "data" / "finance.db"


class FinanceDatabase:

    def __init__(self, db_path: Path = DB_PATH) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(exist_ok=True)
        self._init_schema()

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        return conn

    # ──────────────────────────────────────────────
    # Schema 初始化
    # ──────────────────────────────────────────────

    def _init_schema(self) -> None:
        with self._conn() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS transactions (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    date        TEXT    NOT NULL,
                    description TEXT    NOT NULL,
                    amount      REAL    NOT NULL,
                    category    TEXT    NOT NULL DEFAULT '其他',
                    subcategory TEXT,
                    source      TEXT,
                    import_id   INTEGER,
                    tags        TEXT    DEFAULT '[]',
                    notes       TEXT,
                    is_income   INTEGER NOT NULL DEFAULT 0,
                    created_at  TEXT    DEFAULT (datetime('now')),
                    updated_at  TEXT    DEFAULT (datetime('now'))
                );

                CREATE INDEX IF NOT EXISTS idx_tx_date     ON transactions(date);
                CREATE INDEX IF NOT EXISTS idx_tx_category ON transactions(category);
                CREATE INDEX IF NOT EXISTS idx_tx_amount   ON transactions(amount);

                CREATE TABLE IF NOT EXISTS budgets (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    category     TEXT    NOT NULL,
                    year         INTEGER NOT NULL,
                    month        INTEGER NOT NULL,
                    limit_amount REAL    NOT NULL,
                    UNIQUE(category, year, month)
                );

                CREATE TABLE IF NOT EXISTS imports (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    filename     TEXT,
                    source_hint  TEXT,
                    imported_at  TEXT    DEFAULT (datetime('now')),
                    record_count INTEGER DEFAULT 0,
                    status       TEXT    DEFAULT 'ok'
                );

                CREATE TABLE IF NOT EXISTS settings (
                    key   TEXT PRIMARY KEY,
                    value TEXT
                );

                -- 預設類別設定
                INSERT OR IGNORE INTO settings(key, value) VALUES
                  ('categories', '[
                    {"name":"餐飲","icon":"🍽️","color":"#FF6B6B"},
                    {"name":"交通","icon":"🚌","color":"#4ECDC4"},
                    {"name":"購物","icon":"🛍️","color":"#45B7D1"},
                    {"name":"娛樂","icon":"🎮","color":"#96CEB4"},
                    {"name":"醫療","icon":"🏥","color":"#FFEAA7"},
                    {"name":"教育","icon":"📚","color":"#DDA0DD"},
                    {"name":"住宅","icon":"🏠","color":"#98D8C8"},
                    {"name":"旅遊","icon":"✈️","color":"#FFB347"},
                    {"name":"收入","icon":"💵","color":"#10b981"},
                    {"name":"其他","icon":"📦","color":"#9ca3af"}
                  ]'),
                  ('currency', 'TWD'),
                  ('anthropic_api_key', '');
            """)
        logger.info("[DB] 資料庫初始化完成: %s", self.db_path)

    # ──────────────────────────────────────────────
    # 交易記錄 CRUD
    # ──────────────────────────────────────────────

    def insert_transactions(self, rows: List[Dict[str, Any]], import_id: Optional[int] = None) -> int:
        """批量插入交易記錄，回傳插入筆數"""
        inserted = 0
        with self._conn() as conn:
            for row in rows:
                try:
                    conn.execute(
                        """INSERT INTO transactions
                           (date, description, amount, category, subcategory,
                            source, import_id, tags, notes, is_income)
                           VALUES (?,?,?,?,?,?,?,?,?,?)""",
                        (
                            row.get("date"),
                            row.get("description", ""),
                            row.get("amount", 0),
                            row.get("category", "其他"),
                            row.get("subcategory"),
                            row.get("source"),
                            import_id,
                            json.dumps(row.get("tags", []), ensure_ascii=False),
                            row.get("notes"),
                            1 if row.get("is_income") else 0,
                        ),
                    )
                    inserted += 1
                except sqlite3.IntegrityError:
                    pass
        return inserted

    def get_transactions(
        self,
        start: Optional[str] = None,
        end: Optional[str] = None,
        category: Optional[str] = None,
        search: Optional[str] = None,
        is_income: Optional[int] = None,
        limit: int = 200,
        offset: int = 0,
    ) -> Tuple[List[Dict], int]:
        """查詢交易記錄，回傳 (資料列表, 總筆數)"""
        conditions = ["1=1"]
        params: List[Any] = []

        if start:
            conditions.append("date >= ?")
            params.append(start)
        if end:
            conditions.append("date <= ?")
            params.append(end)
        if category and category != "全部":
            conditions.append("category = ?")
            params.append(category)
        if search:
            conditions.append("description LIKE ?")
            params.append(f"%{search}%")
        if is_income is not None:
            conditions.append("is_income = ?")
            params.append(is_income)

        where = " AND ".join(conditions)

        with self._conn() as conn:
            total = conn.execute(
                f"SELECT COUNT(*) FROM transactions WHERE {where}", params
            ).fetchone()[0]

            rows = conn.execute(
                f"""SELECT * FROM transactions
                    WHERE {where}
                    ORDER BY date DESC, id DESC
                    LIMIT ? OFFSET ?""",
                params + [limit, offset],
            ).fetchall()

        return [dict(r) for r in rows], total

    def get_transaction(self, tx_id: int) -> Optional[Dict]:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT * FROM transactions WHERE id=?", (tx_id,)
            ).fetchone()
        return dict(row) if row else None

    def update_transaction(self, tx_id: int, data: Dict[str, Any]) -> bool:
        allowed = {"date", "description", "amount", "category", "subcategory", "notes", "tags"}
        sets = {k: v for k, v in data.items() if k in allowed}
        if not sets:
            return False
        sets["updated_at"] = datetime.now().isoformat()
        cols = ", ".join(f"{k}=?" for k in sets)
        with self._conn() as conn:
            c = conn.execute(
                f"UPDATE transactions SET {cols} WHERE id=?",
                list(sets.values()) + [tx_id],
            )
        return c.rowcount > 0

    def delete_transaction(self, tx_id: int) -> bool:
        with self._conn() as conn:
            c = conn.execute("DELETE FROM transactions WHERE id=?", (tx_id,))
        return c.rowcount > 0

    def delete_import(self, import_id: int) -> int:
        """刪除某次匯入的所有記錄，回傳刪除筆數"""
        with self._conn() as conn:
            c = conn.execute(
                "DELETE FROM transactions WHERE import_id=?", (import_id,)
            )
            conn.execute("DELETE FROM imports WHERE id=?", (import_id,))
        return c.rowcount

    # ──────────────────────────────────────────────
    # 分析查詢
    # ──────────────────────────────────────────────

    def get_monthly_summary(self, year: int, month: int) -> Dict[str, Any]:
        """取得某月收支摘要"""
        start = f"{year:04d}-{month:02d}-01"
        # 計算月末
        if month == 12:
            end = f"{year+1:04d}-01-01"
        else:
            end = f"{year:04d}-{month+1:02d}-01"

        with self._conn() as conn:
            rows = conn.execute(
                """SELECT
                       SUM(CASE WHEN is_income=0 THEN ABS(amount) ELSE 0 END) as expense,
                       SUM(CASE WHEN is_income=1 THEN amount ELSE 0 END)       as income,
                       COUNT(CASE WHEN is_income=0 THEN 1 END)                  as expense_count,
                       COUNT(CASE WHEN is_income=1 THEN 1 END)                  as income_count
                   FROM transactions
                   WHERE date >= ? AND date < ?""",
                (start, end),
            ).fetchone()

        r = dict(rows) if rows else {}
        return {
            "year":          year,
            "month":         month,
            "expense":       round(r.get("expense") or 0, 2),
            "income":        round(r.get("income") or 0, 2),
            "net":           round((r.get("income") or 0) - (r.get("expense") or 0), 2),
            "expense_count": r.get("expense_count") or 0,
            "income_count":  r.get("income_count") or 0,
        }

    def get_category_breakdown(self, year: int, month: int) -> List[Dict[str, Any]]:
        """取得某月各類別支出分布"""
        start = f"{year:04d}-{month:02d}-01"
        end   = f"{year:04d}-{month+1:02d}-01" if month < 12 else f"{year+1:04d}-01-01"
        with self._conn() as conn:
            rows = conn.execute(
                """SELECT category,
                          SUM(ABS(amount)) as total,
                          COUNT(*) as count
                   FROM transactions
                   WHERE date >= ? AND date < ? AND is_income=0
                   GROUP BY category
                   ORDER BY total DESC""",
                (start, end),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_monthly_trend(self, months: int = 12) -> List[Dict[str, Any]]:
        """取得近 N 個月的收支趨勢"""
        result = []
        today  = datetime.today()
        for i in range(months - 1, -1, -1):
            y = today.year  - (today.month - 1 - i + 11) // 12
            m = (today.month - 1 - i) % 12 + 1
            result.append(self.get_monthly_summary(y, m))
        return result

    def get_daily_expenses(self, start: str, end: str) -> List[Dict[str, Any]]:
        """取得日期範圍內每日支出（用於趨勢線圖）"""
        with self._conn() as conn:
            rows = conn.execute(
                """SELECT date, SUM(ABS(amount)) as total
                   FROM transactions
                   WHERE date >= ? AND date <= ? AND is_income=0
                   GROUP BY date
                   ORDER BY date""",
                (start, end),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_top_merchants(self, year: int, month: int, limit: int = 10) -> List[Dict[str, Any]]:
        """取得某月消費最多的商家"""
        start = f"{year:04d}-{month:02d}-01"
        end   = f"{year:04d}-{month+1:02d}-01" if month < 12 else f"{year+1:04d}-01-01"
        with self._conn() as conn:
            rows = conn.execute(
                """SELECT description, category,
                          SUM(ABS(amount)) as total,
                          COUNT(*) as count
                   FROM transactions
                   WHERE date >= ? AND date < ? AND is_income=0
                   GROUP BY description
                   ORDER BY total DESC
                   LIMIT ?""",
                (start, end, limit),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_recent_transactions(self, limit: int = 10) -> List[Dict[str, Any]]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM transactions ORDER BY date DESC, id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_all_for_ai(self, months: int = 3) -> List[Dict[str, Any]]:
        """供 AI 分析用：近 N 月交易摘要（不含 raw json）"""
        since = (datetime.today() - timedelta(days=months * 31)).strftime("%Y-%m-%d")
        with self._conn() as conn:
            rows = conn.execute(
                """SELECT date, description, amount, category, is_income
                   FROM transactions
                   WHERE date >= ?
                   ORDER BY date DESC
                   LIMIT 500""",
                (since,),
            ).fetchall()
        return [dict(r) for r in rows]

    # ──────────────────────────────────────────────
    # 預算
    # ──────────────────────────────────────────────

    def upsert_budget(self, category: str, year: int, month: int, limit_amount: float) -> int:
        with self._conn() as conn:
            c = conn.execute(
                """INSERT INTO budgets(category, year, month, limit_amount)
                   VALUES(?,?,?,?)
                   ON CONFLICT(category, year, month)
                   DO UPDATE SET limit_amount=excluded.limit_amount""",
                (category, year, month, limit_amount),
            )
        return c.lastrowid

    def get_budgets(self, year: int, month: int) -> List[Dict[str, Any]]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM budgets WHERE year=? AND month=?", (year, month)
            ).fetchall()
        return [dict(r) for r in rows]

    def delete_budget(self, budget_id: int) -> bool:
        with self._conn() as conn:
            c = conn.execute("DELETE FROM budgets WHERE id=?", (budget_id,))
        return c.rowcount > 0

    # ──────────────────────────────────────────────
    # 匯入歷程
    # ──────────────────────────────────────────────

    def create_import(self, filename: str, source_hint: str = "") -> int:
        with self._conn() as conn:
            c = conn.execute(
                "INSERT INTO imports(filename, source_hint) VALUES(?,?)",
                (filename, source_hint),
            )
        return c.lastrowid

    def update_import(self, import_id: int, record_count: int, status: str = "ok") -> None:
        with self._conn() as conn:
            conn.execute(
                "UPDATE imports SET record_count=?, status=? WHERE id=?",
                (record_count, status, import_id),
            )

    def get_imports(self) -> List[Dict[str, Any]]:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM imports ORDER BY imported_at DESC LIMIT 50"
            ).fetchall()
        return [dict(r) for r in rows]

    # ──────────────────────────────────────────────
    # 設定
    # ──────────────────────────────────────────────

    def get_setting(self, key: str) -> Optional[str]:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT value FROM settings WHERE key=?", (key,)
            ).fetchone()
        return row["value"] if row else None

    def set_setting(self, key: str, value: str) -> None:
        with self._conn() as conn:
            conn.execute(
                "INSERT INTO settings(key,value) VALUES(?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )

    def get_all_settings(self) -> Dict[str, str]:
        with self._conn() as conn:
            rows = conn.execute("SELECT key, value FROM settings").fetchall()
        return {r["key"]: r["value"] for r in rows}

    def count_transactions(self) -> int:
        with self._conn() as conn:
            return conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]

    def clear_all_transactions(self) -> int:
        with self._conn() as conn:
            c = conn.execute("DELETE FROM transactions")
            conn.execute("DELETE FROM imports")
        return c.rowcount
