"""
消費分析工具 — FastAPI 後端主程式
啟動後開啟 http://localhost:8765
"""

import json
import logging
import os
import sys
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import uvicorn
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.StreamHandler(sys.stderr)],
)
logger = logging.getLogger(__name__)

from db.database import FinanceDatabase
from core.parser import parse_file
from core.categorizer import categorize_transactions, get_category_for_description

app = FastAPI(title="消費分析工具", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

db = FinanceDatabase()

# ── 靜態檔案 ────────────────────────────────────────────────
STATIC_DIR = PROJECT_ROOT / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


# ══════════════════════════════════════════════════════════
# 交易記錄 API
# ══════════════════════════════════════════════════════════

@app.get("/api/transactions")
async def get_transactions(
    start:     Optional[str] = None,
    end:       Optional[str] = None,
    category:  Optional[str] = None,
    search:    Optional[str] = None,
    is_income: Optional[int] = None,
    limit:     int = Query(default=100, le=500),
    offset:    int = 0,
):
    rows, total = db.get_transactions(
        start=start, end=end, category=category,
        search=search, is_income=is_income,
        limit=limit, offset=offset,
    )
    return {"data": rows, "total": total, "limit": limit, "offset": offset}


class TransactionCreate(BaseModel):
    date:        str
    description: str
    amount:      float
    category:    str = "其他"
    is_income:   int = 0
    notes:       Optional[str] = None


@app.post("/api/transactions", status_code=201)
async def create_transaction(body: TransactionCreate):
    inserted = db.insert_transactions([body.model_dump()])
    if not inserted:
        raise HTTPException(500, "新增失敗")
    rows, _ = db.get_transactions(limit=1, offset=0)
    return rows[0] if rows else {}


class TransactionUpdate(BaseModel):
    date:        Optional[str]   = None
    description: Optional[str]   = None
    amount:      Optional[float] = None
    category:    Optional[str]   = None
    notes:       Optional[str]   = None


@app.put("/api/transactions/{tx_id}")
async def update_transaction(tx_id: int, body: TransactionUpdate):
    ok = db.update_transaction(tx_id, body.model_dump(exclude_none=True))
    if not ok:
        raise HTTPException(404, "找不到該筆記錄")
    return db.get_transaction(tx_id)


@app.delete("/api/transactions/{tx_id}")
async def delete_transaction(tx_id: int):
    ok = db.delete_transaction(tx_id)
    if not ok:
        raise HTTPException(404, "找不到該筆記錄")
    return {"deleted": True}


# ══════════════════════════════════════════════════════════
# 檔案匯入 API
# ══════════════════════════════════════════════════════════

@app.post("/api/import")
async def import_file(file: UploadFile = File(...)):
    """上傳 CSV / Excel 並自動解析、分類、存入資料庫"""
    content  = await file.read()
    filename = file.filename or "upload"

    logger.info("[Import] 收到檔案: %s (%d bytes)", filename, len(content))

    # 建立匯入記錄
    import_id = db.create_import(filename)

    try:
        rows, source = parse_file(content, filename)
        if not rows:
            db.update_import(import_id, 0, "empty")
            return {"success": False, "message": "未解析到有效交易記錄，請確認格式是否正確", "count": 0}

        # 自動分類
        rows = categorize_transactions(rows)

        # 存入資料庫
        inserted = db.insert_transactions(rows, import_id=import_id)
        db.update_import(import_id, inserted, "ok")

        logger.info("[Import] 成功匯入 %d 筆 (來源: %s)", inserted, source)
        return {
            "success":   True,
            "message":   f"成功匯入 {inserted} 筆交易記錄",
            "count":     inserted,
            "source":    source,
            "import_id": import_id,
        }

    except Exception as exc:
        logger.error("[Import] 失敗: %s", exc, exc_info=True)
        db.update_import(import_id, 0, f"error: {exc}")
        raise HTTPException(500, f"解析失敗：{exc}")


@app.get("/api/imports")
async def list_imports():
    return db.get_imports()


@app.delete("/api/imports/{import_id}")
async def delete_import(import_id: int):
    count = db.delete_import(import_id)
    return {"deleted": count}


# ══════════════════════════════════════════════════════════
# 分析 API
# ══════════════════════════════════════════════════════════

@app.get("/api/analysis/summary")
async def get_summary(
    year:  int = Query(default=datetime.today().year),
    month: int = Query(default=datetime.today().month),
):
    current = db.get_monthly_summary(year, month)
    # 上月
    if month == 1:
        prev = db.get_monthly_summary(year - 1, 12)
    else:
        prev = db.get_monthly_summary(year, month - 1)

    # 環比變化
    def pct_change(cur, pre):
        if not pre:
            return None
        return round((cur - pre) / pre * 100, 1)

    return {
        "current":  current,
        "previous": prev,
        "expense_change": pct_change(current["expense"], prev["expense"]),
        "income_change":  pct_change(current["income"],  prev["income"]),
        "total_records":  db.count_transactions(),
    }


@app.get("/api/analysis/categories")
async def get_categories(
    year:  int = Query(default=datetime.today().year),
    month: int = Query(default=datetime.today().month),
):
    breakdown   = db.get_category_breakdown(year, month)
    total_exp   = sum(r["total"] for r in breakdown)
    for r in breakdown:
        r["pct"] = round(r["total"] / total_exp * 100, 1) if total_exp else 0
    return {"breakdown": breakdown, "total_expense": round(total_exp, 2)}


@app.get("/api/analysis/monthly")
async def get_monthly_trend(months: int = Query(default=12, le=24)):
    return db.get_monthly_trend(months)


@app.get("/api/analysis/top-merchants")
async def get_top_merchants(
    year:  int = Query(default=datetime.today().year),
    month: int = Query(default=datetime.today().month),
    limit: int = Query(default=8, le=20),
):
    return db.get_top_merchants(year, month, limit)


@app.get("/api/analysis/recent")
async def get_recent(limit: int = Query(default=10, le=50)):
    return db.get_recent_transactions(limit)


# ══════════════════════════════════════════════════════════
# 預算 API
# ══════════════════════════════════════════════════════════

class BudgetBody(BaseModel):
    category:     str
    year:         int
    month:        int
    limit_amount: float


@app.get("/api/budgets")
async def get_budgets(
    year:  int = Query(default=datetime.today().year),
    month: int = Query(default=datetime.today().month),
):
    budgets   = db.get_budgets(year, month)
    breakdown = db.get_category_breakdown(year, month)
    spent_map = {r["category"]: r["total"] for r in breakdown}

    result = []
    for b in budgets:
        spent = spent_map.get(b["category"], 0)
        pct   = round(spent / b["limit_amount"] * 100, 1) if b["limit_amount"] else 0
        result.append({**b, "spent": round(spent, 2), "pct": pct})

    return result


@app.post("/api/budgets", status_code=201)
async def set_budget(body: BudgetBody):
    bid = db.upsert_budget(body.category, body.year, body.month, body.limit_amount)
    return {"id": bid, **body.model_dump()}


@app.delete("/api/budgets/{budget_id}")
async def delete_budget(budget_id: int):
    ok = db.delete_budget(budget_id)
    if not ok:
        raise HTTPException(404, "找不到預算設定")
    return {"deleted": True}


# ══════════════════════════════════════════════════════════
# AI 分析 API（Server-Sent Events 串流）
# ══════════════════════════════════════════════════════════

class AIAnalysisRequest(BaseModel):
    analysis_type: str = "monthly_review"
    year:          int = datetime.today().year
    month:         int = datetime.today().month


@app.post("/api/ai/analyze")
async def ai_analyze(body: AIAnalysisRequest):
    """串流回傳 AI 分析結果"""
    from ai.claude_analyzer import stream_analysis

    api_key  = db.get_setting("anthropic_api_key") or ""
    tx_data  = db.get_all_for_ai(months=3)
    monthly  = db.get_monthly_summary(body.year, body.month)
    cats     = db.get_category_breakdown(body.year, body.month)
    budgets  = db.get_budgets(body.year, body.month)

    async def sse_generator():
        async for chunk in stream_analysis(
            analysis_type=body.analysis_type,
            transactions=tx_data,
            monthly_summary=monthly,
            category_breakdown=cats,
            budgets=budgets,
            api_key=api_key or None,
        ):
            # SSE format
            yield f"data: {json.dumps({'text': chunk}, ensure_ascii=False)}\n\n"
        yield "data: [DONE]\n\n"

    return StreamingResponse(sse_generator(), media_type="text/event-stream")


# ══════════════════════════════════════════════════════════
# 設定 API
# ══════════════════════════════════════════════════════════

@app.get("/api/settings")
async def get_settings():
    s = db.get_all_settings()
    # 隱藏 API key 明文
    if s.get("anthropic_api_key"):
        s["anthropic_api_key_set"] = True
        s["anthropic_api_key"] = "••••••••••••••••"
    else:
        s["anthropic_api_key_set"] = False
    return s


class SettingUpdate(BaseModel):
    key:   str
    value: str


@app.post("/api/settings")
async def update_setting(body: SettingUpdate):
    db.set_setting(body.key, body.value)
    return {"key": body.key, "updated": True}


@app.delete("/api/data/all")
async def clear_all_data():
    count = db.clear_all_transactions()
    return {"deleted": count}


# ══════════════════════════════════════════════════════════
# 分類建議 API
# ══════════════════════════════════════════════════════════

@app.get("/api/categorize")
async def categorize_hint(description: str, is_income: int = 0):
    cat = get_category_for_description(description, bool(is_income))
    return {"category": cat}


# ══════════════════════════════════════════════════════════
# 啟動
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Windows 終端 UTF-8 相容設定
    if sys.platform == "win32":
        import os
        os.system("chcp 65001 >nul 2>&1")
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    PORT = 8765
    print(f"\n  [消費分析工具] 啟動中...")
    print(f"  伺服器地址: http://localhost:{PORT}")
    print(f"  資料庫: {db.db_path}")
    print(f"  共 {db.count_transactions()} 筆交易記錄")
    print(f"  按 Ctrl+C 停止\n")

    # 延遲開啟瀏覽器
    import threading
    def _open():
        import time; time.sleep(1.5)
        webbrowser.open(f"http://localhost:{PORT}")
    threading.Thread(target=_open, daemon=True).start()

    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="warning")
