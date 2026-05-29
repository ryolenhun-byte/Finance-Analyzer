"""
CSV / Excel / PDF 解析器
自動偵測各銀行、信用卡的匯出格式，轉換為統一的交易記錄格式

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
    - Public Bank Malaysia（大眾銀行馬來西亞）逐字座標解析
    - 台灣各銀行存摺/明細 PDF（含表格的 PDF）
    - 信用卡帳單 PDF
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
# PDF 解析入口
# ══════════════════════════════════════════════

def _parse_pdf(content: bytes, filename: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    解析銀行/信用卡 PDF 帳單

    優先策略：
      0. 偵測是否為 Public Bank Malaysia 格式 → 專用解析器
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

            # ── 策略 0a：UOB Malaysia 專用解析器（優先於 PBB，避免 PENYATA AKAUN 誤判）
            if _is_uob_pdf(pdf, source):
                uob_rows, uob_source = _parse_uob_pdf(pdf, filename)
                if uob_rows:
                    logger.info("[PDF] UOB 模式：提取 %d 筆", len(uob_rows))
                    return uob_rows, uob_source
                logger.warning("[PDF] UOB 模式無結果，嘗試通用模式")

            # ── 策略 0b：Public Bank Malaysia 專用解析器 ────────────────
            if _is_public_bank_pdf(pdf, source):
                source = "Public Bank Malaysia"
                rows = _parse_public_bank_pdf(pdf, filename)
                if rows:
                    logger.info("[PDF] Public Bank 模式：提取 %d 筆", len(rows))
                    return rows, source
                logger.warning("[PDF] Public Bank 模式無結果，嘗試通用模式")

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
    # 馬來西亞銀行
    if any(k in fname for k in ["public bank", "publicbank", "pbb", "pb bank"]):
        return "Public Bank Malaysia"
    if any(k in fname for k in ["uob", "united overseas"]):
        return "UOB Malaysia(PDF)"
    if "maybank" in fname or "mbb" in fname:
        return "Maybank(PDF)"
    if "cimb" in fname:
        return "CIMB(PDF)"
    if "rhb" in fname:
        return "RHB(PDF)"
    if "hlb" in fname or "hong leong" in fname:
        return "Hong Leong Bank(PDF)"
    # 台灣銀行
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


def _is_public_bank_pdf(pdf, source: str) -> bool:
    """
    判斷是否為 Public Bank Malaysia 格式。
    必須明確包含 "PUBLIC BANK" 或 "PBB" 才確認，
    避免與其他馬來西亞銀行（UOB 等）衝突。
    """
    if "public bank" in source.lower():
        return True
    try:
        first_page = pdf.pages[0]
        text = (first_page.extract_text() or "").upper()
        # 必須明確出現 "PUBLIC BANK" 或 "PBB" 才算（UOB/Maybank 不包含這些字）
        if "PUBLIC BANK" in text or " PBB " in text:
            return True
        # 備用：DUITNOW + BAKI（PBB 特有組合，且不含 UOB 關鍵字）
        if "DUITNOW" in text and "BAKI" in text and "UOB" not in text:
            return True
    except Exception:
        pass
    return False


# ══════════════════════════════════════════════
# Public Bank Malaysia 專用解析器
# ══════════════════════════════════════════════

# X 座標分欄界線（根據對 Public Bank PDF 的分析）
_PB_X_DATE_END    = 88   # 日期欄：x < 88
_PB_X_DESC_END    = 310  # 描述欄：88 <= x < 310
_PB_X_DEBIT_END   = 390  # 借方欄：310 <= x < 390
_PB_X_CREDIT_END  = 465  # 貸方欄：390 <= x < 465
# 餘額欄：x >= 465

# 頁首/頁尾行關鍵字，用來跳過非交易行
_PB_HEADER_SKIP = {
    "tarikh", "transaksi", "debit", "kredit", "baki",
    "date", "transaction", "credit", "balance",
    "penyata akaun", "account statement",
    "nama", "name", "no.", "akaun", "account",
    "cawangan", "branch", "kod", "code",
}

# 描述中需要過濾掉的浮水印 / 行政字詞（馬來語 PDF 常見）
_PB_DESC_FILTER_WORDS = {
    "dicetak", "melalui", "komputer",   # "printed via computer"
    "printed", "computer",
    "penyata",                           # "statement"
    "diperbuat",                         # "made"
}

# 需要整行過濾的描述片段（出現即跳過整個 desc 行）
_PB_DESC_FILTER_PHRASES = [
    "dicetak melalui",
    "printed via",
    "this is a computer",
]

# 不算作交易的特殊行標記
_PB_NON_TX_MARKERS = [
    "balance b/f", "baki b/h", "baki b/f",
    "balance from last", "balance from previous",
    "brought forward", "carried forward",
    "balance c/f", "baki c/h",
    "closing balance", "opening balance",
]


def _parse_public_bank_pdf(pdf, filename: str) -> List[Dict[str, Any]]:
    """
    解析 Public Bank Malaysia（大眾銀行馬來西亞）帳單 PDF

    使用 extract_words() 的 x 座標來識別欄位：
      - 日期 (DATE):        x < 88      格式 DD/MM
      - 描述 (DESCRIPTION): 88 <= x < 310
      - 借方 (DEBIT):       310 <= x < 390
      - 貸方 (CREDIT):      390 <= x < 465
      - 餘額 (BALANCE):     x >= 465

    多行描述：第一行有日期，後續行無日期但有描述文字。
    """
    source = "Public Bank Malaysia"
    stmt_year, stmt_month = _extract_pb_date_from_filename(filename)
    rows: List[Dict[str, Any]] = []
    current_tx: Optional[Dict] = None

    for page_num, page in enumerate(pdf.pages, 1):
        # 嘗試從頁首提取年月（更準確）
        year_from_header = _extract_pb_year_from_page(page)
        if year_from_header:
            stmt_year = year_from_header

        words = page.extract_words(
            x_tolerance=3,
            y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=False,
        )
        if not words:
            continue

        # 按 y 座標分組成行
        lines = _pb_group_words_by_y(words, y_tolerance=5)

        for line_words in lines:
            if not line_words:
                continue

            # 按欄位分類
            date_parts, desc_parts, debit_parts, credit_parts = [], [], [], []

            for w in line_words:
                x    = float(w.get("x0", 0))
                text = w.get("text", "").strip()
                if not text:
                    continue

                if x < _PB_X_DATE_END:
                    date_parts.append(text)
                elif x < _PB_X_DESC_END:
                    desc_parts.append(text)
                elif x < _PB_X_DEBIT_END:
                    debit_parts.append(text)
                elif x < _PB_X_CREDIT_END:
                    credit_parts.append(text)
                # 餘額欄忽略

            date_str   = " ".join(date_parts).strip()
            desc_str   = " ".join(desc_parts).strip()
            debit_str  = " ".join(debit_parts).strip()
            credit_str = " ".join(credit_parts).strip()

            # 跳過純表頭行（描述為單個表頭關鍵字）
            if _is_pb_header_line(date_str, desc_str):
                continue

            # 判斷是否是新交易行（有 DD/MM 日期）
            date_match = re.match(r"^(\d{1,2})/(\d{2})$", date_str)

            if date_match:
                # 儲存上一筆
                if current_tx and _pb_is_valid_tx(current_tx):
                    tx = _pb_finalize_tx(current_tx, stmt_year, stmt_month, source)
                    if tx:
                        rows.append(tx)

                # 檢查是否為 Balance B/F 等非交易行
                full_text = (date_str + " " + desc_str).lower()
                if any(marker in full_text for marker in _PB_NON_TX_MARKERS):
                    current_tx = None
                    continue

                # 開始新交易
                current_tx = {
                    "date_raw":   date_str,
                    "desc_lines": [desc_str] if desc_str else [],
                    "debit":      debit_str,
                    "credit":     credit_str,
                }

            elif desc_str and current_tx is not None:
                # 續行：追加描述
                desc_lower = desc_str.lower()
                # 跳過表頭型文字
                if not any(hk in desc_lower for hk in _PB_HEADER_SKIP):
                    current_tx["desc_lines"].append(desc_str)
                # 若首行沒取到金額，嘗試從續行補充
                if debit_str and not current_tx["debit"]:
                    current_tx["debit"] = debit_str
                if credit_str and not current_tx["credit"]:
                    current_tx["credit"] = credit_str

        # 頁末：不在這裡 flush，讓交易跨頁延續（Public Bank 說明可能跨頁）

    # 最後一筆
    if current_tx and _pb_is_valid_tx(current_tx):
        tx = _pb_finalize_tx(current_tx, stmt_year, stmt_month, source)
        if tx:
            rows.append(tx)

    logger.info("[Public Bank] 共解析 %d 筆交易（%d 年 %d 月帳單）",
                len(rows), stmt_year, stmt_month)
    return rows


def _pb_group_words_by_y(words: List[Dict], y_tolerance: int = 5) -> List[List[Dict]]:
    """按 y 座標（top）將文字分組為行，同行容許差異 y_tolerance 個單位"""
    if not words:
        return []

    sorted_words = sorted(words, key=lambda w: (float(w.get("top", 0)), float(w.get("x0", 0))))
    lines = []
    current_line = [sorted_words[0]]
    current_y    = float(sorted_words[0].get("top", 0))

    for word in sorted_words[1:]:
        wy = float(word.get("top", 0))
        if abs(wy - current_y) <= y_tolerance:
            current_line.append(word)
        else:
            lines.append(sorted(current_line, key=lambda w: float(w.get("x0", 0))))
            current_line = [word]
            current_y    = wy

    if current_line:
        lines.append(sorted(current_line, key=lambda w: float(w.get("x0", 0))))

    return lines


def _extract_pb_date_from_filename(filename: str) -> Tuple[int, int]:
    """
    從檔名推斷年月
    支援格式：
      - "Jan 2026.pdf"  →  (2026, 1)
      - "2026-01.pdf"   →  (2026, 1)
      - "202601.pdf"    →  (2026, 1)
    """
    MONTH_MAP = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4,
        "may": 5, "jun": 6, "jul": 7, "aug": 8,
        "sep": 9, "oct": 10, "nov": 11, "dec": 12,
        "january": 1, "february": 2, "march": 3, "april": 4,
        "june": 6, "july": 7, "august": 8, "september": 9,
        "october": 10, "november": 11, "december": 12,
    }

    stem = Path(filename).stem.lower().strip()

    # 年份
    year_m = re.search(r"\b(20\d{2})\b", stem)
    year   = int(year_m.group(1)) if year_m else datetime.today().year

    # 月份（英文月份名）
    month = datetime.today().month
    for name, num in MONTH_MAP.items():
        if name in stem:
            month = num
            break
    else:
        # 數字月份
        num_m = re.search(r"\b(0?[1-9]|1[0-2])\b", stem)
        if num_m:
            month = int(num_m.group(1))

    return year, month


def _extract_pb_year_from_page(page) -> Optional[int]:
    """
    從頁面文字提取帳單年份
    尋找 "Statement Date DD Mon YYYY" 或 "Tarikh Penyata DD Mon YYYY"
    """
    try:
        text = page.extract_text() or ""
        # 例："Statement Date 31 Jan 2026" 或 "31 January 2026"
        m = re.search(
            r"(?:statement\s+date|tarikh\s+penyata)[^\d]*(\d{1,2})\s+\w+\s+(20\d{2})",
            text, re.IGNORECASE
        )
        if m:
            return int(m.group(2))
        # 備用：找 "20XX" 在頁面前 500 字元
        m2 = re.search(r"\b(20\d{2})\b", text[:500])
        if m2:
            return int(m2.group(1))
    except Exception:
        pass
    return None


def _is_pb_header_line(date_str: str, desc_str: str) -> bool:
    """判斷是否為表頭/頁首說明行（非交易）"""
    # 若日期欄是表頭文字（非日期數字格式）
    if date_str and not re.match(r"^\d{1,2}/\d{2}$", date_str):
        combined = (date_str + " " + desc_str).lower()
        if any(hk in combined for hk in _PB_HEADER_SKIP):
            return True
    return False


def _pb_is_valid_tx(tx: Dict) -> bool:
    """交易必須有借方或貸方金額"""
    return bool(
        (tx.get("debit") and tx["debit"].strip()) or
        (tx.get("credit") and tx["credit"].strip())
    )


def _pb_clean_amount(s: str) -> Optional[float]:
    """清理 Public Bank 金額字串（移除逗號、空白），回傳 float 或 None"""
    s = re.sub(r"[,\s]", "", s.strip())
    if not s or s in ("-", "—", "N/A"):
        return None
    try:
        return abs(float(s))
    except ValueError:
        return None


def _pb_build_description(desc_lines: List[str]) -> str:
    """
    從多行描述建立乾淨的交易描述

    過濾規則：
    - 跳過超長純數字參考碼（≥ 10 位數字）
    - 跳過 IMEPS... 等銀行內部流水號
    - 跳過馬來語浮水印文字（dicetak melalui kompu 等）
    - 最多保留 80 個字元
    """
    tokens = []
    for line in desc_lines:
        line_lower = line.lower()
        # 跳過整行含浮水印短語
        if any(phrase in line_lower for phrase in _PB_DESC_FILTER_PHRASES):
            continue

        for word in line.split():
            word_lower = word.lower()
            # 跳過超長純數字（參考/流水號）
            if re.match(r"^\d{10,}$", word):
                continue
            # 跳過 IMEPS 內部流水號
            if re.match(r"^imeps\d{8,}", word_lower):
                continue
            # 跳過浮水印關鍵字
            if word_lower in _PB_DESC_FILTER_WORDS:
                continue
            tokens.append(word)

    desc = " ".join(tokens).strip()
    return desc[:80] if len(desc) > 80 else desc


def _pb_finalize_tx(
    tx: Dict, stmt_year: int, stmt_month: int, source: str
) -> Optional[Dict[str, Any]]:
    """將內部交易 dict 轉換為標準格式"""
    # ── 日期 ────────────────────────────────────────────
    date_raw = tx.get("date_raw", "")
    try:
        day, month = map(int, date_raw.split("/"))
        # 年份推斷：若交易月份比帳單月份大超過 1，可能是上年底交易
        if month > stmt_month and (month - stmt_month) > 1:
            year = stmt_year - 1
        else:
            year = stmt_year
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
    except Exception:
        date_str = f"{stmt_year:04d}-{stmt_month:02d}-01"

    # ── 描述 ────────────────────────────────────────────
    desc = _pb_build_description(tx.get("desc_lines", []))
    if not desc:
        desc = "Public Bank Transaction"

    # ── 金額 ────────────────────────────────────────────
    debit_val  = _pb_clean_amount(tx.get("debit", ""))
    credit_val = _pb_clean_amount(tx.get("credit", ""))

    if credit_val and credit_val > 0:
        amount    = credit_val
        is_income = 1
    elif debit_val and debit_val > 0:
        amount    = -debit_val
        is_income = 0
    else:
        return None  # 無有效金額，丟棄

    return {
        "date":        date_str,
        "description": desc,
        "amount":      round(amount, 2),
        "is_income":   is_income,
        "source":      source,
        "category":    "其他",
    }


# ══════════════════════════════════════════════
# UOB Malaysia 專用解析器
# ══════════════════════════════════════════════

# UOB 月份縮寫對應表（帳單常見大寫縮寫）
_UOB_MONTH_MAP = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,
    "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,
    "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

# UOB Savings — X 座標分欄（依 extract_words 實際觀測值）
_UOB_SAV_X_TRANS_END = 90    # Trans Date: x < 90  (day + month)
_UOB_SAV_X_VDATE_END = 145   # Value Date: 90 <= x < 145
_UOB_SAV_X_DESC_END  = 330   # Description: 145 <= x < 330
_UOB_SAV_X_WD_END    = 410   # Withdrawals: 330 <= x < 410
_UOB_SAV_X_DEP_END   = 490   # Deposits:    410 <= x < 490
# Balance: x >= 490

# UOB Credit Card — X 座標分欄
_UOB_CC_X_DATE_END   = 120   # Date: x < 120  (day + month)
_UOB_CC_X_DESC_END   = 460   # Description: 120 <= x < 460
# Amount: x >= 460  (with optional "CR" suffix)

# 跳過的非交易行關鍵字（大寫比對）
_UOB_SAV_SKIP_DESC = {
    "BALANCE B/F", "BAKI B/H", "BAKI B/F",
    "BALANCE C/F", "BAKI C/H",
    "BALANCE B/F",
}
_UOB_SAV_SKIP_PARTIAL = [
    "balance b/f", "baki b/", "closing balance", "opening balance",
]

# UOB CC — 跳過的行（bc interest / bc instalment / retail interest 屬真實費用，不跳過）
_UOB_CC_SKIP_PARTIAL = [
    "previous bal", "credit limit", "sub-total", "minimum payment",
    "end of statement",
    # 頁尾地址行
    "united overseas bank", "menara uob", "jalan raja laut",
]

# UOB CC — 新卡段標頭識別（含 ** 的卡號行）
_UOB_CC_CARD_HEADER_RE = re.compile(r"\*\*[\d\-]+\*\*")


def _is_uob_pdf(pdf, source: str) -> bool:
    """判斷是否為 UOB Malaysia PDF（儲蓄帳戶或信用卡）"""
    if "uob" in source.lower():
        return True
    try:
        first_page = pdf.pages[0]
        text = (first_page.extract_text() or "").upper()
        # UOB 特有識別：郵件地址、卡片中心名稱、ONE Account
        uob_markers = [
            "UOBCUSTOMERSERVICE@UOB.COM.MY",
            "UOB CARD CENTRE",
            "UNITED OVERSEAS BANK",
            "ONE ACCOUNT",
        ]
        return any(kw in text for kw in uob_markers)
    except Exception:
        return False


def _parse_uob_pdf(pdf, filename: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    自動偵測並分派 UOB 儲蓄帳戶 或 UOB 信用卡 解析器

    Returns:
        (rows, source_name)
    """
    try:
        first_text = (pdf.pages[0].extract_text() or "").upper()
    except Exception:
        first_text = ""

    # 信用卡：必須明確出現 "UOB CARD CENTRE"（信用卡帳單獨有）
    # 儲蓄帳戶：含 "ONE ACCOUNT" 或 "SIMPANAN"，不含 "CARD CENTRE"
    is_credit_card = "UOB CARD CENTRE" in first_text
    if is_credit_card:
        rows = _parse_uob_credit_card(pdf, filename)
        return rows, "UOB Malaysia Credit Card"
    else:
        rows = _parse_uob_savings(pdf, filename)
        return rows, "UOB Malaysia ONE Account"


# ─── UOB 儲蓄帳戶解析器 ─────────────────────────────────────────────────────

def _parse_uob_savings(pdf, filename: str) -> List[Dict[str, Any]]:
    """
    解析 UOB Malaysia ONE Account 月結單

    版面特徵：
      - 第 1-2 頁為帳戶總覽，第 3 頁起為交易明細
      - 頁首含 "Account Transaction Details" / "Butiran Transaksi Akaun"
      - 欄位（x 座標）：
        Trans Date < 90 | Value Date 90-145 | Description 145-330
        Withdrawals 330-410 | Deposits 410-490 | Balance ≥ 490
      - 日期格式：DD Mon（例 "07 Feb"）
      - 跨年處理：Period 行已包含完整年份
    """
    rows: List[Dict[str, Any]] = []
    current_tx: Optional[Dict] = None
    stmt_year, stmt_month = _uob_extract_period(pdf)
    source = "UOB Malaysia ONE Account"

    for page_num, page in enumerate(pdf.pages, 1):
        text = page.extract_text() or ""
        # 只處理含交易明細的頁面
        if not any(kw in text for kw in [
            "Account Transaction Details", "Butiran Transaksi Akaun",
            "BALANCE B/F", "BAKI B/H",
        ]):
            continue

        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        if not words:
            continue

        lines = _pb_group_words_by_y(words, y_tolerance=6)

        for line_words in lines:
            if not line_words:
                continue

            # 分欄
            date_parts, desc_parts, wd_parts, dep_parts = [], [], [], []
            for w in line_words:
                x    = float(w.get("x0", 0))
                text = w.get("text", "").strip()
                if not text:
                    continue
                if x < _UOB_SAV_X_TRANS_END:
                    date_parts.append(text)
                elif x < _UOB_SAV_X_VDATE_END:
                    pass  # 忽略 Value Date
                elif x < _UOB_SAV_X_DESC_END:
                    desc_parts.append(text)
                elif x < _UOB_SAV_X_WD_END:
                    wd_parts.append(text)
                elif x < _UOB_SAV_X_DEP_END:
                    dep_parts.append(text)
                # balance 忽略

            date_str = " ".join(date_parts).strip()
            desc_str = " ".join(desc_parts).strip()
            wd_str   = " ".join(wd_parts).strip()
            dep_str  = " ".join(dep_parts).strip()

            # 跳過表頭行
            if _uob_is_header_line(date_str, desc_str):
                continue

            # 是否為新交易行（Trans Date = "DD Mon" 格式）
            date_match = re.match(
                r"^(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)$",
                date_str, re.IGNORECASE
            )

            if date_match:
                # 儲存上一筆
                if current_tx and _uob_sav_is_valid(current_tx):
                    tx = _uob_sav_finalize(current_tx, stmt_year, stmt_month, source)
                    if tx:
                        rows.append(tx)

                # 跳過非交易行
                desc_lower = desc_str.lower()
                if any(m in desc_lower for m in _UOB_SAV_SKIP_PARTIAL):
                    current_tx = None
                    continue

                current_tx = {
                    "date_raw":   date_str,
                    "desc_lines": [desc_str] if desc_str else [],
                    "wd":         wd_str,
                    "dep":        dep_str,
                }

            elif desc_str and current_tx is not None:
                # 續行：補充描述
                desc_lower = desc_str.lower()
                skip_words = {"trans", "date", "value", "description",
                              "withdrawals", "deposits", "balance",
                              "tarikh", "transaksi", "deskripsi"}
                if desc_str.lower() not in skip_words:
                    # 過濾純分隔符號
                    if desc_str not in ("-", "|", "/", "\\"):
                        current_tx["desc_lines"].append(desc_str)
                # 補充未取到的金額
                if wd_str and not current_tx["wd"]:
                    current_tx["wd"] = wd_str
                if dep_str and not current_tx["dep"]:
                    current_tx["dep"] = dep_str

    # 最後一筆
    if current_tx and _uob_sav_is_valid(current_tx):
        tx = _uob_sav_finalize(current_tx, stmt_year, stmt_month, source)
        if tx:
            rows.append(tx)

    logger.info("[UOB Savings] 解析 %d 筆（%d 年 %d 月）", len(rows), stmt_year, stmt_month)
    return rows


def _uob_extract_period(pdf) -> Tuple[int, int]:
    """從 UOB PDF 的 Period 行提取年月"""
    MONTH_MAP = _UOB_MONTH_MAP
    for page in pdf.pages[:2]:
        text = page.extract_text() or ""
        # "Period: 01 Feb 2026 to 28 Feb 2026"
        m = re.search(
            r"Period\s*:\s*\d{1,2}\s+(\w{3})\s+(\d{4})",
            text, re.IGNORECASE
        )
        if m:
            mon_str = m.group(1).upper()
            year    = int(m.group(2))
            month   = MONTH_MAP.get(mon_str, datetime.today().month)
            return year, month
        # "Tempoh: 01 Feb 2026 sehingga ..."
        m2 = re.search(
            r"Tempoh\s*:\s*\d{1,2}\s+(\w{3})\s+(\d{4})",
            text, re.IGNORECASE
        )
        if m2:
            mon_str = m2.group(1).upper()
            year    = int(m2.group(2))
            month   = MONTH_MAP.get(mon_str, datetime.today().month)
            return year, month
    return datetime.today().year, datetime.today().month


def _uob_is_header_line(date_str: str, desc_str: str) -> bool:
    """判斷是否為表頭/頁首行"""
    header_words = {
        "trans", "date", "value", "description", "withdrawals",
        "deposits", "balance", "rm", "tarikh", "transaksi",
        "deskripsi", "pengeluaran", "deposit", "baki",
        "account", "transaction", "details", "butiran",
    }
    if not date_str and desc_str:
        if desc_str.lower() in header_words:
            return True
    if date_str and not re.match(r"^\d{1,2}\s+\w{3}$", date_str, re.IGNORECASE):
        combined = (date_str + " " + desc_str).lower()
        if any(hw in combined for hw in header_words):
            return True
    return False


def _uob_sav_is_valid(tx: Dict) -> bool:
    return bool(
        (tx.get("wd") and tx["wd"].strip()) or
        (tx.get("dep") and tx["dep"].strip())
    )


def _uob_clean_amount(s: str) -> Optional[float]:
    """清理 UOB 金額字串"""
    s = re.sub(r"[,\s]", "", s.strip())
    if not s or s in ("-", "—"):
        return None
    try:
        return abs(float(s))
    except ValueError:
        return None


def _uob_sav_finalize(
    tx: Dict, stmt_year: int, stmt_month: int, source: str
) -> Optional[Dict[str, Any]]:
    """將 UOB Savings 內部交易 dict 轉換為標準格式"""
    # 日期
    date_raw = tx.get("date_raw", "")
    try:
        parts   = date_raw.split()
        day     = int(parts[0])
        month   = _UOB_MONTH_MAP.get(parts[1].upper(), 0)
        if not month:
            return None
        # 年份推斷：若交易月份比帳單月份大超過 1，為上年底交易
        if month > stmt_month and (month - stmt_month) > 1:
            year = stmt_year - 1
        else:
            year = stmt_year
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
    except Exception:
        date_str = f"{stmt_year:04d}-{stmt_month:02d}-01"

    # 描述
    desc_lines = tx.get("desc_lines", [])
    desc = " ".join(desc_lines).strip()
    # 過濾純數字流水號（16位帳號等）
    tokens = []
    for token in desc.split():
        if re.match(r"^\d{10,}$", token):
            continue
        tokens.append(token)
    desc = " ".join(tokens).strip()
    if not desc:
        desc = "UOB Transaction"
    desc = desc[:80]

    # 金額
    wd_val  = _uob_clean_amount(tx.get("wd", ""))
    dep_val = _uob_clean_amount(tx.get("dep", ""))

    if dep_val and dep_val > 0:
        amount    = dep_val
        is_income = 1
    elif wd_val and wd_val > 0:
        amount    = -wd_val
        is_income = 0
    else:
        return None

    return {
        "date":        date_str,
        "description": desc,
        "amount":      round(amount, 2),
        "is_income":   is_income,
        "source":      source,
        "category":    "其他",
    }


# ─── UOB 信用卡解析器 ──────────────────────────────────────────────────────

def _parse_uob_credit_card(pdf, filename: str) -> List[Dict[str, Any]]:
    """
    解析 UOB Malaysia 信用卡月結單（UOB CARD CENTRE STATEMENT OF ACCOUNT）

    版面特徵：
      - 前幾頁為說明頁，交易從 "STATEMENT / PENYATA" 頁開始
      - 一份帳單可含多張卡（VISA INFINITE / PRVI MILES 等），以卡號行分隔
      - 每張卡有 "PREVIOUS BAL xxx" 後才是交易
      - 欄位（x 座標）：
        Date < 120 (DD MON) | Description 120-460 | Amount >= 460
        若末尾有 "CR"，表示貸方（退款/還款/退稅）
      - 帳單年份從 "Statement Date DD MON YY" 取得
      - Skip: PREVIOUS BAL, CREDIT LIMIT, SUB-TOTAL, MINIMUM PAYMENT, RETAIL INTEREST
             BC INTEREST, BC INSTALMENT, ANNUAL FEE, FOREIGN TRX (額外手續費行)
    """
    rows: List[Dict[str, Any]] = []
    source = "UOB Malaysia Credit Card"

    # 從第一頁取得帳單年月
    stmt_year, stmt_month = _uob_cc_extract_stmt_date(pdf)

    # 哪些頁面含交易資料？
    for page_num, page in enumerate(pdf.pages, 1):
        text = page.extract_text() or ""
        # 跳過不含交易的頁面（說明/法律頁）
        if not (
            "Transaction Date" in text or "Tarikh Transaksi" in text
            or re.search(r"\b\d{2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b", text)
        ):
            continue

        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        if not words:
            continue

        lines = _pb_group_words_by_y(words, y_tolerance=6)
        current_tx: Optional[Dict] = None

        for line_words in lines:
            if not line_words:
                continue

            date_parts, desc_parts, amt_parts = [], [], []
            for w in line_words:
                x    = float(w.get("x0", 0))
                wt   = w.get("text", "").strip()
                if not wt:
                    continue
                if x < _UOB_CC_X_DATE_END:
                    date_parts.append(wt)
                elif x < _UOB_CC_X_DESC_END:
                    desc_parts.append(wt)
                else:
                    amt_parts.append(wt)

            date_str = " ".join(date_parts).strip()
            desc_str = " ".join(desc_parts).strip()
            amt_str  = " ".join(amt_parts).strip()

            # 跳過表頭行
            if _uob_cc_is_header(date_str, desc_str):
                continue

            # 跳過卡段標頭行（含 "**xxxx-xxxx**" 的行，代表新卡開始）
            if _UOB_CC_CARD_HEADER_RE.search(desc_str):
                # 儲存上一筆
                if current_tx:
                    tx = _uob_cc_finalize(current_tx, stmt_year, stmt_month, source)
                    if tx:
                        rows.append(tx)
                    current_tx = None
                continue

            # 跳過非交易行
            desc_lower = (date_str + " " + desc_str).lower()
            if any(m in desc_lower for m in _UOB_CC_SKIP_PARTIAL):
                if current_tx:
                    tx = _uob_cc_finalize(current_tx, stmt_year, stmt_month, source)
                    if tx:
                        rows.append(tx)
                    current_tx = None
                continue

            # 新交易行：DD MON 格式
            date_match = re.match(
                r"^(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)$",
                date_str, re.IGNORECASE
            )

            if date_match:
                # 儲存上一筆
                if current_tx:
                    tx = _uob_cc_finalize(current_tx, stmt_year, stmt_month, source)
                    if tx:
                        rows.append(tx)

                current_tx = {
                    "date_raw":   date_str,
                    "desc_lines": [desc_str] if desc_str else [],
                    "amt_raw":    amt_str,
                }

            elif desc_str and current_tx is not None:
                # 續行：追加描述（跳過純幣別行如 "USD" 或純數字金額行）
                if re.match(r"^[A-Z]{3}$", desc_str):
                    # 幣別行，忽略
                    pass
                elif not amt_str and re.match(r"^[\d,\.]+$", desc_str):
                    # 外幣金額續行（如 "200.00"），忽略
                    pass
                else:
                    current_tx["desc_lines"].append(desc_str)
                # 補充金額
                if amt_str and not current_tx["amt_raw"]:
                    current_tx["amt_raw"] = amt_str

        # 頁末 flush
        if current_tx:
            tx = _uob_cc_finalize(current_tx, stmt_year, stmt_month, source)
            if tx:
                rows.append(tx)
            current_tx = None

    logger.info("[UOB CC] 解析 %d 筆（帳單日 %d 年 %d 月）", len(rows), stmt_year, stmt_month)
    return rows


def _uob_cc_extract_stmt_date(pdf) -> Tuple[int, int]:
    """從信用卡帳單取得帳單年月（Statement Date DD MON YY）"""
    for page in pdf.pages[:2]:
        text = page.extract_text() or ""
        # "Statement Date 05 MAY 26" 或 "Statement Date 05 MAY 2026"
        m = re.search(
            r"Statement\s+Date\s+\d{1,2}\s+(\w{3})\s+(\d{2,4})",
            text, re.IGNORECASE
        )
        if m:
            mon_str  = m.group(1).upper()
            yr_raw   = m.group(2)
            year     = int(yr_raw) + (2000 if len(yr_raw) == 2 else 0)
            month    = _UOB_MONTH_MAP.get(mon_str, datetime.today().month)
            return year, month
    return datetime.today().year, datetime.today().month


def _uob_cc_is_header(date_str: str, desc_str: str) -> bool:
    """判斷是否為 UOB CC 表頭行"""
    header_words = {
        "transaction", "date", "description", "amount",
        "tarikh", "transaksi", "huraian", "amaun", "rm",
        "statement", "penyata",
    }
    combined = (date_str + " " + desc_str).lower().split()
    if not date_str and all(w in header_words for w in combined if w):
        return True
    if date_str and not re.match(r"^\d{1,2}\s+\w{3}$", date_str, re.IGNORECASE):
        if any(hw in (date_str + " " + desc_str).lower() for hw in header_words):
            return True
    return False


def _uob_cc_finalize(
    tx: Dict, stmt_year: int, stmt_month: int, source: str
) -> Optional[Dict[str, Any]]:
    """將 UOB CC 內部交易 dict 轉換為標準格式"""
    # 日期
    date_raw = tx.get("date_raw", "")
    try:
        parts  = date_raw.split()
        day    = int(parts[0])
        month  = _UOB_MONTH_MAP.get(parts[1].upper(), 0)
        if not month:
            return None
        # 年份推斷：信用卡帳單可跨月（e.g. 05 MAY 帳單含 APR 交易）
        # 若交易月份 > 帳單月份+1 → 上年
        if month > stmt_month and (month - stmt_month) > 1:
            year = stmt_year - 1
        else:
            year = stmt_year
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
    except Exception:
        date_str = f"{stmt_year:04d}-{stmt_month:02d}-01"

    # 描述
    desc_lines = tx.get("desc_lines", [])
    desc = " ".join(desc_lines).strip()
    # 移除城市/國家代碼後綴（如 "KUALA LUMPUR MY"）
    desc = re.sub(r"\s+[A-Z]{2}$", "", desc).strip()
    if not desc:
        desc = "UOB CC Transaction"
    desc = desc[:80]

    # 金額 & 方向
    amt_raw = tx.get("amt_raw", "").strip()
    is_credit = amt_raw.endswith("CR")
    amt_clean = re.sub(r"[,\s]", "", amt_raw.replace("CR", "").strip())
    try:
        val = float(amt_clean)
    except ValueError:
        return None
    if val == 0:
        return None

    if is_credit:
        # CR = 還款/退款 → 收入
        amount    = val
        is_income = 1
    else:
        # 消費
        amount    = -val
        is_income = 0

    return {
        "date":        date_str,
        "description": desc,
        "amount":      round(amount, 2),
        "is_income":   is_income,
        "source":      source,
        "category":    "其他",
    }


# ══════════════════════════════════════════════
# 通用 PDF 解析（既有邏輯）
# ══════════════════════════════════════════════

def _extract_pdf_tables(pdf, source: str) -> List[Dict[str, Any]]:
    """從 PDF 表格中提取交易（最準確的方式）"""
    all_rows: List[Dict[str, Any]] = []

    for page_num, page in enumerate(pdf.pages, 1):
        tables = page.extract_tables()
        for table in tables:
            if not table or len(table) < 2:
                continue

            header_idx, col_map = _find_table_header(table)
            if header_idx is None or not col_map.get("date"):
                continue

            for row in table[header_idx + 1:]:
                if not row or all(not str(c).strip() for c in row if c is not None):
                    continue
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

    for i, row in enumerate(table[:5]):
        if not row:
            continue
        cells   = [str(c).strip() if c else "" for c in row]
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

    _find(_DATE_COLS,    "date")
    _find(_DESC_COLS,    "description")
    _find(_DEBIT_COLS,   "debit")
    _find(_CREDIT_COLS,  "credit")
    _find(_BALANCE_COLS, "balance")
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

    amount, is_income = _parse_amount_strings(get("debit"), get("credit"))
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
    適用於沒有結構化表格的 PDF
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

    date_pattern   = re.compile(
        r"(\d{4}[/\-]\d{1,2}[/\-]\d{1,2}|\d{1,2}/\d{1,2}(?:/\d{2,4})?)"
    )
    amount_pattern = re.compile(r"-?[\d,]+(?:\.\d+)?")

    for line in lines:
        line = line.strip()
        if len(line) < 5:
            continue

        date_match = date_pattern.search(line)
        if not date_match:
            continue

        date = _parse_date(date_match.group(0))
        if not date:
            continue

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

        after_date      = line[date_match.end():].strip()
        first_num_match = re.search(r"-?[\d,]+", after_date)
        if first_num_match:
            desc = after_date[:first_num_match.start()].strip()
        else:
            desc = after_date[:30].strip()

        if not desc:
            desc = "PDF 交易"

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
    """從一行的多個數字中，選出最可能是「交易金額」的那個"""
    if not nums:
        return None
    if len(nums) == 1:
        return nums[0]
    if len(nums) >= 2:
        for v in nums:
            if 1 <= abs(v) <= 500_000:
                return -abs(v)
    return nums[0]


# ══════════════════════════════════════════════
# CSV / Excel 讀取器
# ══════════════════════════════════════════════

def _read_csv(content: bytes, filename: str) -> Tuple[Optional[pd.DataFrame], str]:
    """嘗試多種編碼讀取 CSV"""
    detected  = chardet.detect(content)
    encodings = [detected.get("encoding") or "utf-8", "utf-8", "big5", "cp950", "utf-8-sig"]
    encodings = list(dict.fromkeys(e for e in encodings if e))

    for enc in encodings:
        try:
            text = content.decode(enc)
            text = text.lstrip("﻿")
            df   = pd.read_csv(io.StringIO(text), dtype=str, on_bad_lines="skip")
            df   = df.dropna(how="all")
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
    cols_lower  = {c.lower() for c in columns}
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
    date_part = raw[:10].strip()
    for fmt in formats:
        try:
            dt = datetime.strptime(date_part, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

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
        s = re.sub(r"[,\s$NT$TWD元RM]", "", s)
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
        s = re.sub(r"[,\s$NT$TWD元RM]", "", s.strip())
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
