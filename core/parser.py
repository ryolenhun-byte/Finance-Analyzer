"""
CSV / Excel 解析器
自動偵測台灣各銀行、信用卡的匯出格式，轉換為統一的交易記錄格式

支援格式：
  - 玉山銀行 (E.SUN)
  - 國泰世華 (Cathay United)
  - 中信銀行 (CTBC)
  - 台新銀行 (Taishin)
  - 一般 CSV（日期/描述/金額）
  - Excel (.xlsx / .xls)
"""

import io
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import chardet
import pandas as pd

logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# 欄位名稱模式（優先順序由高到低）
# ──────────────────────────────────────────────

_DATE_COLS = [
    "交易日期", "日期", "消費日期", "帳務日期", "記帳日期",
    "交易日期時間", "date", "Date", "交易日",
]
_DESC_COLS = [
    "摘要", "說明", "消費說明", "交易說明", "交易摘要", "備註",
    "商店名稱", "店名", "消費店名", "description", "Description",
    "交易描述", "對象", "交易對象",
]
_DEBIT_COLS = [
    "支出金額", "支出", "消費金額", "交易金額", "借方金額",
    "扣款金額", "金額", "debit", "Debit", "amount", "Amount",
    "出帳金額",
]
_CREDIT_COLS = [
    "收入金額", "收入", "存入金額", "貸方金額", "入帳金額",
    "credit", "Credit",
]
_BALANCE_COLS = [
    "餘額", "可用餘額", "帳戶餘額", "balance", "Balance",
]


def parse_file(content: bytes, filename: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    解析上傳的檔案，回傳 (交易列表, 偵測到的來源名稱)

    Args:
        content  : 檔案原始位元組
        filename : 原始檔名

    Returns:
        (rows, source_hint) — rows 為統一格式的交易字典列表
    """
    ext = Path(filename).suffix.lower()

    if ext in (".xlsx", ".xls"):
        df, source = _read_excel(content, filename)
    else:
        df, source = _read_csv(content, filename)

    if df is None or df.empty:
        return [], source

    rows = _normalize_dataframe(df, source)
    logger.info("[Parser] %s: 解析 %d 筆交易", filename, len(rows))
    return rows, source


# ──────────────────────────────────────────────
# 讀取器
# ──────────────────────────────────────────────

def _read_csv(content: bytes, filename: str) -> Tuple[Optional[pd.DataFrame], str]:
    """嘗試多種編碼讀取 CSV"""
    # 自動偵測編碼
    detected = chardet.detect(content)
    encodings = [detected.get("encoding") or "utf-8", "utf-8", "big5", "cp950", "utf-8-sig"]
    encodings = list(dict.fromkeys(e for e in encodings if e))  # 去重保序

    for enc in encodings:
        try:
            text = content.decode(enc)
            # 移除 BOM
            text = text.lstrip("﻿")
            df = pd.read_csv(io.StringIO(text), dtype=str, on_bad_lines="skip")
            # 移除全空列
            df = df.dropna(how="all")
            if not df.empty:
                source = _detect_source(df.columns.tolist(), filename)
                logger.debug("[Parser] CSV 編碼: %s，來源: %s", enc, source)
                return df, source
        except Exception as exc:
            logger.debug("[Parser] 嘗試編碼 %s 失敗: %s", enc, exc)

    return None, "未知"


def _read_excel(content: bytes, filename: str) -> Tuple[Optional[pd.DataFrame], str]:
    """讀取 Excel 檔案"""
    try:
        df = pd.read_excel(io.BytesIO(content), dtype=str, header=0)
        df = df.dropna(how="all")
        if not df.empty:
            source = _detect_source(df.columns.tolist(), filename)
            return df, source
    except Exception as exc:
        logger.error("[Parser] Excel 讀取失敗: %s", exc)
    return None, "未知"


# ──────────────────────────────────────────────
# 來源偵測
# ──────────────────────────────────────────────

def _detect_source(columns: List[str], filename: str) -> str:
    """根據欄位名稱或檔名判斷資料來源"""
    cols_lower = {c.lower() for c in columns}
    fname_lower = filename.lower()

    # 根據欄位特徵識別
    if "可用餘額" in cols_lower or "esun" in fname_lower:
        return "玉山銀行"
    if "帳務日期" in cols_lower or "cathay" in fname_lower:
        return "國泰世華"
    if "借方金額" in cols_lower or "ctbc" in fname_lower:
        return "中信銀行"
    if "taishin" in fname_lower or "台新" in fname_lower:
        return "台新銀行"
    if "line" in fname_lower:
        return "Line Pay"
    if "jko" in fname_lower or "街口" in fname_lower:
        return "街口支付"
    return "一般 CSV"


# ──────────────────────────────────────────────
# 標準化
# ──────────────────────────────────────────────

def _normalize_dataframe(df: pd.DataFrame, source: str) -> List[Dict[str, Any]]:
    """將 DataFrame 欄位對應到統一格式"""
    cols = df.columns.tolist()
    col_map = _build_column_map(cols)

    if not col_map.get("date") or not col_map.get("description"):
        logger.warning("[Parser] 無法辨識必要欄位，欄位列表: %s", cols)
        return []

    rows: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        try:
            date = _parse_date(str(row.get(col_map["date"], "") or ""))
            if not date:
                continue

            desc = str(row.get(col_map["description"], "") or "").strip()
            if not desc or desc.lower() in ("nan", "none", ""):
                continue

            amount, is_income = _parse_amount(row, col_map)
            if amount is None:
                continue

            rows.append({
                "date":        date,
                "description": desc,
                "amount":      amount,
                "is_income":   is_income,
                "source":      source,
                "category":    "其他",  # 後續由 categorizer 處理
            })

        except Exception as exc:
            logger.debug("[Parser] 跳過列: %s", exc)

    return rows


def _build_column_map(columns: List[str]) -> Dict[str, str]:
    """將已知欄位名稱映射到標準欄位 key"""
    col_lower_map = {c.strip(): c for c in columns}  # 清理空白
    cols_set = set(col_lower_map.keys())

    def _find(candidates: List[str]) -> Optional[str]:
        for c in candidates:
            if c in cols_set:
                return c
            # 模糊匹配（欄位名包含關鍵字）
            for col in cols_set:
                if c in col or col in c:
                    return col
        return None

    return {
        "date":        _find(_DATE_COLS),
        "description": _find(_DESC_COLS),
        "debit":       _find(_DEBIT_COLS),
        "credit":      _find(_CREDIT_COLS),
        "balance":     _find(_BALANCE_COLS),
    }


def _parse_date(raw: str) -> Optional[str]:
    """解析多種日期格式，回傳 YYYY-MM-DD"""
    raw = raw.strip()
    if not raw or raw.lower() in ("nan", "none"):
        return None

    # 嘗試多種格式
    formats = [
        "%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y",
        "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S",
        "%Y%m%d",
    ]
    for fmt in formats:
        try:
            # 只取日期部分（忽略時間）
            dt = datetime.strptime(raw[:len(fmt.replace("%H:%M:%S", "").strip())].strip(),
                                   fmt.split(" ")[0])
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # 試用 pandas 解析
    try:
        return pd.to_datetime(raw).strftime("%Y-%m-%d")
    except Exception:
        return None


def _parse_amount(row: pd.Series, col_map: Dict[str, str]) -> Tuple[Optional[float], int]:
    """
    解析金額，回傳 (amount, is_income)
    支出為負數（is_income=0），收入為正數（is_income=1）
    """
    debit_col  = col_map.get("debit")
    credit_col = col_map.get("credit")

    def _clean(val: Any) -> Optional[float]:
        if val is None:
            return None
        s = str(val).strip()
        if not s or s.lower() in ("nan", "none", ""):
            return None
        # 移除貨幣符號、逗號、空白
        s = re.sub(r"[,\s$NT$TWD元]", "", s)
        try:
            return float(s)
        except ValueError:
            return None

    # 有分開的支出/收入欄位
    if debit_col and credit_col:
        debit  = _clean(row.get(debit_col))
        credit = _clean(row.get(credit_col))
        if credit and credit > 0:
            return credit, 1
        if debit and debit != 0:
            return -abs(debit), 0
        return None, 0

    # 只有金額欄（正=收入，負=支出，或根據欄位判斷）
    if debit_col:
        val = _clean(row.get(debit_col))
        if val is None:
            return None, 0
        if val > 0:
            # 有些格式收入也在同欄，用正號表示
            return -val, 0   # 預設視為支出
        elif val < 0:
            return val, 0    # 已帶負號
        return None, 0

    return None, 0
