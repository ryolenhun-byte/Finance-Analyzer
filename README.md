# 💰 消費分析工具

個人消費記錄分析 Web 應用，支援匯入銀行 CSV/Excel、自動分類、趨勢圖表與 Claude AI 深度分析。

## 功能

| 功能 | 說明 |
|------|------|
| 📤 多格式匯入 | 玉山、國泰、中信、台新銀行 CSV / Excel，自動偵測格式 |
| 🏷️ 自動分類 | 關鍵字規則自動分類（餐飲/交通/購物/娛樂等） |
| 📊 儀表板 | 月度摘要、類別圓餅圖、6個月趨勢柱狀圖 |
| 🎯 預算管理 | 各類別每月預算上限，進度條顯示使用率 |
| 🤖 AI 分析 | Claude AI 深度分析（本月綜合/消費習慣/省錢建議/異常偵測） |
| 🔄 持續累積 | SQLite 本機儲存，每次匯入自動累積，可隨時查詢歷史 |

## 快速開始

```bash
# 安裝依賴
pip install -r requirements.txt

# 啟動（自動開啟瀏覽器）
python main.py
```

或直接雙擊 `启动.bat`

瀏覽器開啟：**http://localhost:8765**

## 匯入格式

支援的銀行/平台：
- 玉山銀行 (E.SUN)
- 國泰世華 (Cathay United)
- 中信銀行 (CTBC)
- 台新銀行 (Taishin)
- Line Pay、街口支付
- 一般 CSV（日期/描述/金額）

## AI 分析（可選）

需要 Anthropic API Key，在設定頁面配置。

分析類型：
- 📊 本月綜合分析
- 🔄 消費習慣分析
- 💡 省錢建議
- ⚠️ 異常消費偵測

## 技術棧

- **後端**：Python FastAPI + SQLite
- **前端**：Tailwind CSS + Chart.js（無需 Node.js）
- **AI**：Claude claude-sonnet-4-6（串流輸出 + Prompt Caching）
