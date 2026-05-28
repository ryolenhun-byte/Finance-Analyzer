"""
Claude AI 消費分析模組
使用 Claude claude-sonnet-4-6 進行深度財務分析，支援串流輸出與 Prompt Caching
"""

import json
import logging
import os
from typing import Any, AsyncGenerator, Dict, List, Optional

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """你是一位專業的個人理財顧問，專精於分析台灣消費者的消費習慣。

你的分析風格：
- 直接、具體、有實際操作價值
- 使用繁體中文，語氣親切但專業
- 數字分析要精確，給出百分比和金額
- 建議要可行，考慮台灣的生活水平與物價
- 避免空泛的「建議節省開銷」，要指出具體的項目和方法

回應格式使用 Markdown，適度使用 emoji 讓報告更易讀。"""


async def stream_analysis(
    analysis_type: str,
    transactions: List[Dict[str, Any]],
    monthly_summary: Optional[Dict] = None,
    category_breakdown: Optional[List] = None,
    budgets: Optional[List] = None,
    api_key: Optional[str] = None,
) -> AsyncGenerator[str, None]:
    """
    串流回傳 AI 分析結果

    Args:
        analysis_type : monthly_review | habits | saving_tips | anomaly
        transactions  : 近期交易列表
        monthly_summary: 月度摘要
        category_breakdown: 類別分布
        budgets       : 預算設定
        api_key       : Anthropic API Key

    Yields:
        文字片段
    """
    try:
        import anthropic
    except ImportError:
        yield "❌ 請先安裝 anthropic 套件：pip install anthropic"
        return

    key = api_key or os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        yield "❌ 未設定 Anthropic API Key，請前往「設定」頁面配置。"
        return

    prompt = _build_prompt(
        analysis_type, transactions, monthly_summary, category_breakdown, budgets
    )

    client = anthropic.Anthropic(api_key=key)

    try:
        with client.messages.stream(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            system=[
                {
                    "type": "text",
                    "text": SYSTEM_PROMPT,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            for chunk in stream.text_stream:
                yield chunk

    except anthropic.AuthenticationError:
        yield "\n\n❌ API Key 無效或已過期，請重新設定。"
    except anthropic.APIConnectionError:
        yield "\n\n❌ 網路連線失敗，請檢查網路後重試。"
    except anthropic.RateLimitError:
        yield "\n\n❌ 請求頻率超限，請稍後再試。"
    except Exception as exc:
        logger.error("[AI] 分析失敗: %s", exc, exc_info=True)
        yield f"\n\n❌ 分析失敗：{exc}"


def _build_prompt(
    analysis_type: str,
    transactions: List[Dict],
    monthly_summary: Optional[Dict],
    category_breakdown: Optional[List],
    budgets: Optional[List],
) -> str:
    """根據分析類型組合 prompt"""

    # 壓縮交易資料（限制 token 用量）
    tx_summary = _summarize_transactions(transactions)
    budget_str = _format_budgets(budgets or [])
    monthly_str = _format_monthly(monthly_summary)
    category_str = _format_categories(category_breakdown or [])

    prompts = {
        "monthly_review": f"""請對以下消費數據進行本月度綜合分析：

{monthly_str}

## 各類別支出
{category_str}

## 主要交易記錄（近期）
{tx_summary}

{budget_str}

請分析：
1. 📊 **本月消費總結** — 整體支出水平評估
2. 🔍 **類別深度分析** — 哪些類別超出正常水平？哪些控制良好？
3. 📈 **與預算比較** — 哪些類別超預算或接近預算？
4. 💡 **本月重點建議** — 3 條最有價值的改善建議""",

        "habits": f"""請深度分析以下消費模式，找出消費習慣的規律與問題：

## 交易記錄
{tx_summary}

## 類別分布
{category_str}

請分析：
1. 🔄 **消費規律** — 找出固定的消費模式（時間規律、商家偏好等）
2. ⚡ **衝動消費特徵** — 識別可能的非必要消費
3. 🎯 **消費優先級評估** — 支出是否符合生活價值觀？
4. 🛠️ **習慣改善計畫** — 具體可執行的消費習慣調整建議""",

        "saving_tips": f"""基於以下消費數據，提供最具體的省錢建議：

{monthly_str}

## 各類別支出
{category_str}

## 主要交易
{tx_summary}

請提供：
1. 💰 **立即可執行的省錢方法**（每項附估計每月節省金額）
2. 🔄 **訂閱服務審查** — 哪些可能重複或不必要？
3. 🛒 **購物優化建議** — 如何用更少錢達到同樣效果？
4. 📅 **中期理財建議** — 3-6 個月內可執行的財務改善計畫
5. 🎯 **目標儲蓄方案** — 如何每月多存 10-20%？""",

        "anomaly": f"""請仔細審查以下交易記錄，識別異常消費：

## 全部交易記錄
{tx_summary}

請分析：
1. ⚠️ **異常大額消費** — 明顯偏高的單筆消費
2. 🔁 **重複扣款偵測** — 可能的重複收費或意外訂閱
3. ❓ **不明消費** — 描述不清晰、無法識別來源的交易
4. 📉 **消費異常月份** — 某月支出異常增高的原因分析
5. 🔒 **安全建議** — 是否有需要追蹤或爭議的交易？""",
    }

    return prompts.get(
        analysis_type,
        f"請分析以下消費數據並提供改善建議：\n\n{tx_summary}\n\n{category_str}"
    )


def _summarize_transactions(transactions: List[Dict]) -> str:
    """壓縮交易列表為簡潔文字"""
    if not transactions:
        return "（無交易記錄）"

    lines = []
    for tx in transactions[:100]:  # 最多 100 筆
        date     = tx.get("date", "")[:10]
        desc     = tx.get("description", "")[:20]
        amount   = tx.get("amount", 0)
        category = tx.get("category", "其他")
        sign     = "+" if tx.get("is_income") else ""
        lines.append(f"{date}  {desc:<20}  {category:<6}  {sign}{amount:,.0f}")

    return "\n".join(lines)


def _format_monthly(summary: Optional[Dict]) -> str:
    if not summary:
        return ""
    return (
        f"- 本月支出：NT$ {summary.get('expense', 0):,.0f}\n"
        f"- 本月收入：NT$ {summary.get('income', 0):,.0f}\n"
        f"- 淨結餘：NT$ {summary.get('net', 0):,.0f}\n"
        f"- 消費筆數：{summary.get('expense_count', 0)} 筆"
    )


def _format_categories(breakdown: List[Dict]) -> str:
    if not breakdown:
        return "（無分類數據）"
    lines = []
    for item in breakdown:
        lines.append(
            f"- {item.get('category', '?')}: "
            f"NT$ {item.get('total', 0):,.0f} "
            f"（{item.get('count', 0)} 筆）"
        )
    return "\n".join(lines)


def _format_budgets(budgets: List[Dict]) -> str:
    if not budgets:
        return ""
    lines = ["## 預算設定"]
    for b in budgets:
        lines.append(f"- {b.get('category')}: NT$ {b.get('limit_amount', 0):,.0f}")
    return "\n".join(lines)
