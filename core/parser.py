"""
CSV / Excel / PDF 解析器
自動偵測台灣各銀行、信用卡的匯出格式，轉換為統一的交易記錄格式

支援格式：
  CSV / Excel：
    - 玉山銀行 (E.SUN)
    - 國泰世華 (Cathay United)
    - 中信銀行 (CTBC)
    - 台新銀行 (Taishin)
    - Line Pay / 街口支付
    - 一般 CSV（日期/描述/金額）
    - Excel (.xlsx / .xls)
  PDF：
    - 台灣各銀行存摺/明細 PDF（含表格的 PDF）
    - 信用卡帳單 PDF
    - 掃描文字可辨識的 PDF
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

    if ext == ".pdf":
        rows, source = _parse_pdf(content, filename)
    elif ext in (".xlsx", ".xls"):
        df, source = _read_excel(content, filename)
        rows = _normalize_dataframe(df, source) if df is not None else []
    else:
        df, source = _read_csv(content, filename)
        rows = _normalize_dataframe(df, source) if df is not None else []

    logger.info("[Parser] %s: 解析 %d 筆交易", filename, len(rows))
    return rows, source


# ══════════════════════════════════════════════
# PDF 解析
# ══════════════════════════════════════════════

def _parse_pdf(content: bytes, filename: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    解析銀行/信用卡 PDF 帳單

    策略：
      1. 嘗試用 pdfplumber 提取表格（結構化資料）
      2. 若無表格，改用文字行解析（非結構化）
    """
    try:
        import pdfplumber
    except ImportError:
        logger.error("[PDF] pdfplumber 未安裝，請執行: pip install pdfplumber")
        return [], "PDF（需要 pdfplumber）"

    source = _detect_pdf_source(filename)
    rows: List[Dict[str, Any]] = []

    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total_pages = len(pdf.pages)
            logger.info("[PDF] 共 %d 頁，來源識別: %s", total_pages, source)

            # ── 策略 1：逐頁提取表格 ────────────────────────────
            table_rows = _extract_pdf_tables(pdf, source)
            if table_rows:
                logger.info("[PDF] 表格模式：提取 %d 筆", len(table_rows))
                return table_rows, source

            # ── 策略 2：文字行解析 ───────────────────────────────
            text_rows = _extract_pdf_text(pdf, source)
            if text_rows:
                logger.info("[PDF] 文字模式：提取 %d 筆", len(text_rows))
                return text_rows, source

            logger.warning("[PDF] 無法從此 PDF 提取交易資料（可能是掃描圖片）")
            return [], source

    except Exception as exc:
        logger.error("[PDF] 解析失敗: %s", exc, exc_info=True)
        return [], source


def _detect_pdf_source(filename: str) -> str:
    """根據檔名判斷 PDF 來源"""
    fname = filename.lower()
    if "esun" in fname or "玉山" in fname:
        return "玉山銀行(PDF)"
    if "cathay" in fname or "國泰" in fname:
        return "國泰世華(PDF)"
    if "ctbc" in fname or "中信" in fname:
        return "中信銀行(PDF)"
    if "taishin" in fname or "台新" in fname:
        return "台新銀行(PDF)"
    if "sinopac" in fname or "永豐" in fname:
        return "永豐銀行(PDF)"
    if "line" in fname:
        return "Line Pay(PDF)"
    return "銀行帳單(PDF)"


def _extract_pdf_tables(pdf, source: str) -> List[Dict[str, Any]]:
    """從 PDF 表格中提取交易（最準確的方式）"""
    all_rows: List[Dict[str, Any]] = []

    for page_num, page in enumerate(pdf.pages, 1):
        tables = page.extract_tables()
        for table in tables:
            if not table or len(table) < 2:
                continue

            # 找表頭（第一行或前幾行）
            header_idx, col_map = _find_table_header(table)
            if header_idx is None or not col_map.get("date"):
                continue

            # 解析資料列
            for row in table[header_idx + 1:]:
                if not row or all(not str(c).strip() for c in row if c is not None):
                    continue  # 跳過空列
                try:
                    tx = _parse_table_row(row, col_map, source)
                    if tx:
                        all_rows.append(tx)
                except Exception as exc:
                    logger.debug("[PDF表格] 跳過列: %s", exc)

    return all_rows


def _find_table_header(table: List[List]) -> Tuple[Optional[int], Dict[str, int]]:
    """在表格中找到表頭列，回傳 (表頭行index, 欄位映射)"""
    header_keywords = set(
        _DATE_COLS + _DESC_COLS + _DEBIT_COLS + _CREDIT_COLS + _BALANCE_COLS
    )

    for i, row in enumerate(table[:5]):  # 只搜尋前 5 列
        if not row:
            continue
        cells = [str(c).strip() if c else "" for c in row]
        matches = sum(1 for c in cells if any(kw in c for kw in header_keywords))
        if matches >= 2:
            col_map = _build_column_map_from_list(cells)
            if col_map.get("date"):
                return i, col_map

    return None, {}


def _build_column_map_from_list(headers: List[str]) -> Dict[str, int]:
    """從表頭列建立欄位名稱→索引映射"""
    result: Dict[str, int] = {}

    def _find(candidates, key):
        for i, h in enumerate(headers):
            h_clean = h.strip()
            for c in candidates:
                if c in h_clean or h_clean in c:
                    if key not in result:
                        result[key] = i
                    return

    _find(_DATE_COLS,   "date")
    _find(_DESC_COLS,   "description")
    _find(_DEBIT_COLS,  "debit")
    _find(_CREDIT_COLS, "credit")
    _find(_BALANCE_COLS,"balance")
    return result


def _parse_table_row(
    row: List, col_map: Dict[str, int], source: str
) -> Optional[Dict[str, Any]]:
    """解析表格中的一列資料"""
    def get(key: str) -> str:
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return ""
        val = row[idx]
        return str(val).strip() if val is not None else ""

    date = _parse_date(get("date"))
    if not date:
        return None

    desc = get("description")
    if not desc or desc.lower() in ("nan", "none", ""):
        return None

    # 解析金額
    debit_str  = get("debit")
    credit_str = get("credit")

    amount, is_income = _parse_amount_strings(debit_str, credit_str)
    if amount is None:
        return None

    return {
        "date":        date,
        "description": desc,
        "amount":      amount,
        "is_income":   is_income,
        "source":      source,
        "category":    "其他",
    }


def _extract_pdf_text(pdf, source: str) -> List[Dict[str, Any]]:
    """
    從 PDF 純文字中解析交易（備用方案）

    適用於沒有結構化表格的 PDF，
    例如條列式帳單：
      2025/05/01  統一超商  -85  49,915
    """
    full_text = ""
    for page in pdf.pages:
        text = page.extract_text(x_tolerance=2, y_tolerance=2)
        if text:
            full_text += text + "\n"

    if not full_text.strip():
        return []

    rows: List[Dict[str, Any]] = []
    lines = full_text.split("\n")

    # 日期模式（多種格式）
    date_pattern = re.compile(
        r"(\d{4}[/\-]\d{1,2}[/\-]\d{1,2}|\d{1,2}/\d{1,2}(?:/\d{2,4})?)"
    )
    # 金額模式（含千分位逗號，可帶負號）
    amount_pattern = re.compile(r"-?[\d,]+(?:\.\d+)?")

    for line in lines:
        line = line.strip()
        if len(line) < 5:
            continue

        # 找日期
        date_match = date_pattern.search(line)
        if not date_match:
            continue

        date = _parse_date(date_match.group(0))
        if not date:
            continue

        # 從行中提取所有數字
        amounts = amount_pattern.findall(line.replace(",", ""))
        numeric = []
        for a in amounts:
            try:
                v = float(a.replace(",", ""))
                if abs(v) > 0:
                    numeric.append(v)
            except ValueError:
                pass

        if not numeric:
            continue

        # 描述 = 日期之後、第一個數字之前的文字
        after_date = line[date_match.end():].strip()
        first_num_match = re.search(r"-?[\d,]+", after_date)
        if first_num_match:
            desc = after_date[:first_num_match.start()].strip()
        else:
            desc = after_date[:30].strip()

        if not desc:
            desc = "PDF 交易"

        # 取第一個非餘額的金額（通常餘額是最後一個數字且很大）
        amount_val = numeric[0] if len(numeric) == 1 else _pick_transaction_amount(numeric)
        if amount_val is None:
            continue

        is_income = 1 if amount_val > 0 else 0
        rows.append({
            "date":        date,
            "description": desc[:50],
            "amount":      amount_val if is_income else -abs(amount_val),
            "is_income":   is_income,
            "source":      source,
            "category":    "其他",
        })

    return rows


def _pick_transaction_amount(nums: List[float]) -> Optional[float]:
    """
    從一行的多個數字中，選出最可能是「交易金額」的那個
    通常：最後一個大數字是餘額，倒數第二個是金額
    """
    if not nums:
        return None
    if len(nums) == 1:
        return nums[0]
    # 排除異常大的餘額（> 10 倍中位數）
    if len(nums) >= 2:
        for v in nums:
            if 1 <= abs(v) <= 500_000:
                return -abs(v)  # 預設視為支出
    return nums[0]


# ══════════════════════════════════════════════
# CSV / Excel 讀取器（原有邏輯）
# ══════════════════════════════════════════════

def _read_csv(content: bytes, filename: str) -> Tuple[Optional[pd.DataFrame], str]:
    """嘗試多種編碼讀取 CSV"""
    detected = chardet.detect(content)
    encodings = [detected.get("encoding") or "utf-8", "utf-8", "big5", "cp950", "utf-8-sig"]
    encodings = list(dict.fromkeys(e for e in encodings if e))

    for enc in encodings:
        try:
            text = content.decode(enc)
            text = text.lstrip("﻿")
            df = pd.read_csv(io.StringIO(text), dtype=str, on_bad_lines="skip")
            df = df.dropna(how="all")
            if not df.empty:
                source = _detect_source(df.columns.tolist(), filename)
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


def _detect_source(columns: List[str], filename: str) -> str:
    """根據欄位名稱或檔名判斷資料來源"""
    cols_lower = {c.lower() for c in columns}
    fname_lower = filename.lower()

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


# ══════════════════════════════════════════════
# 共用標準化邏輯
# ══════════════════════════════════════════════

def _normalize_dataframe(df: pd.DataFrame, source: str) -> List[Dict[str, Any]]:
    """將 DataFrame 欄位對應到統一格式"""
    if df is None or df.empty:
        return []

    cols    = df.columns.tolist()
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
                "category":    "其他",
            })

        except Exception as exc:
            logger.debug("[Parser] 跳過列: %s", exc)

    return rows


def _build_column_map(columns: List[str]) -> Dict[str, str]:
    """將欄位名稱映射到標準欄位 key"""
    cols_set = set(columns)

    def _find(candidates: List[str]) -> Optional[str]:
        for c in candidates:
            if c in cols_set:
                return c
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


# ══════════════════════════════════════════════
# 共用工具函式
# ══════════════════════════════════════════════

def _parse_date(raw: str) -> Optional[str]:
    """解析多種日期格式，回傳 YYYY-MM-DD"""
    raw = str(raw).strip()
    if not raw or raw.lower() in ("nan", "none"):
        return None

    formats = [
        "%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y",
        "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S",
        "%Y%m%d", "%d/%m/%Y", "%m/%d/%y",
    ]
    # 只取前 10 個字元作為日期部分
    date_part = raw[:10].strip()
    for fmt in formats:
        try:
            dt = datetime.strptime(date_part, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # 備用：pandas 解析
    try:
        return pd.to_datetime(raw).strftime("%Y-%m-%d")
    except Exception:
        return None


def _parse_amount(row: pd.Series, col_map: Dict[str, str]) -> Tuple[Optional[float], int]:
    """
    從 DataFrame 列解析金額，回傳 (amount, is_income)
    支出回傳負數，收入回傳正數
    """
    debit_col  = col_map.get("debit")
    credit_col = col_map.get("credit")

    def _clean(val: Any) -> Optional[float]:
        if val is None:
            return None
        s = str(val).strip()
        if not s or s.lower() in ("nan", "none", ""):
            return None
        s = re.sub(r"[,\s$NT$TWD元]", "", s)
        try:
            return float(s)
        except ValueError:
            return None

    if debit_col and credit_col:
        debit  = _clean(row.get(debit_col))
        credit = _clean(row.get(credit_col))
        if credit and credit > 0:
            return credit, 1
        if debit and debit != 0:
            return -abs(debit), 0
        return None, 0

    if debit_col:
        val = _clean(row.get(debit_col))
        if val is None:
            return None, 0
        if val > 0:
            return -val, 0
        elif val < 0:
            return val, 0
        return None, 0

    return None, 0


def _parse_amount_strings(
    debit_str: str, credit_str: str
) -> Tuple[Optional[float], int]:
    """從兩個字串解析支出/收入金額（PDF 表格用）"""
    def _clean(s: str) -> Optional[float]:
        s = re.sub(r"[,\s$NT$TWD元]", "", s.strip())
        if not s or s in ("-", "—", ""):
            return None
        try:
            return abs(float(s))
        except ValueError:
            return None

    credit = _clean(credit_str)
    if credit and credit > 0:
        return credit, 1

    debit = _clean(debit_str)
    if debit and debit > 0:
        return -debit, 0

    return None, 0
