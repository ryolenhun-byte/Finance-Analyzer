"""
交易自動分類器
使用關鍵字規則快速分類，涵蓋台灣及馬來西亞常見消費場景
"""

import logging
import re
from typing import List, Dict, Any

logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# 類別關鍵字規則（優先順序由高到低）
# ──────────────────────────────────────────────

_RULES: List[Dict[str, Any]] = [
    # ── 收入 ──────────────────────────────────────────────────────────────
    {
        "category": "收入",
        "keywords": [
            # 通用
            "薪資", "薪水", "工資", "獎金", "津貼", "補助", "退款", "退費",
            "利息收入", "dividend", "salary", "payroll",
            "匯入", "存入", "收款",
            # 馬來西亞
            "gaji", "bonus", "elaun", "pulangan", "dividen",
            "trsf cr",          # Public Bank DuitNow 轉入
            "duitnow trsf cr",  # DuitNow 收款
            "ibft cr",          # 網路銀行轉帳收入
            "interbank cr",
            "fpx cr",
            "transfer in",
            "credit transfer",
            "refund",
        ],
        "is_income": True,
    },

    # ── 貸款（高優先，避免被其他關鍵字誤匹配）─────────────────────────────
    {
        "category": "貸款",
        "keywords": [
            "automated loan pymt", "loan pymt", "loan payment",
            "loan repayment", "bayaran pinjaman",
            "home loan", "car loan", "personal loan",
            "pembiayaan", "pinjaman",
            "instalment", "installment",
        ],
    },

    # ── 信用卡還款（高優先）───────────────────────────────────────────────
    {
        "category": "信用卡還款",
        "keywords": [
            "cr card pymt", "credit card pymt", "credit card payment",
            "card payment", "card repayment",
            "cc payment", "cc pymt",
            "bayaran kad kredit",
        ],
    },

    # ── 信用卡費用（高優先，避免被其他規則誤分）──────────────────────────
    {
        "category": "信用卡費用",
        "keywords": [
            # UOB 分期付款（BC = Balance Conversion）
            "bc instalment", "bc installment",
            # 分期利息 & 循環利息
            "bc interest", "retail interest",
            # 年費 & 服務稅（主卡/附屬卡）
            "annual fee", "annual fee waiver",
            "service tax-principal", "service tax-supplementary",
            "service tax",          # 通用服務稅
            "finance charge",
        ],
    },

    # ── 轉帳（高優先，避免轉帳備注文字被誤分為餐飲/娛樂）──────────────────
    {
        "category": "轉帳",
        "keywords": [
            # 馬來西亞常見轉帳描述（越精確越好，避免誤匹配）
            "duitnow trsf dr",      # DuitNow 轉出（Public Bank）
            "duitnow/instant trf",  # DuitNow Instant Transfer（UOB）
            "duitnow instant",
            "duitnow transfer",
            "ibft dr",           # 網路銀行轉帳轉出
            "ibg dr",            # 銀行轉帳轉出
            "tsfr fund dr",      # 資金轉帳轉出
            "interbank transfer",
            "fpx payment",
            "wise pymt",         # Wise 國際匯款（精確版本）
            "wise payment",
            "remittance",
            "western union",
            "transfer out",
            "jompay",            # JomPAY 繳費
        ],
    },

    # ── 餐飲 ──────────────────────────────────────────────────────────────
    {
        "category": "餐飲",
        "keywords": [
            # 台灣
            "超商", "便利商店", "7-eleven", "7-11", "全家", "family mart", "萊爾富", "ok mart",
            "麥當勞", "mcdonald", "肯德基", "kfc", "漢堡王", "burger king",
            "星巴克", "starbucks", "85度c", "路易莎", "coco", "清心",
            "foodpanda", "ubereats", "uber eats", "熊貓外送",
            "餐廳", "食堂", "餐飲", "小吃", "便當", "火鍋", "燒肉",
            "拉麵", "壽司", "迴轉壽司", "日式", "韓式", "泰式",
            "飲料", "咖啡", "茶飲", "珍珠奶茶", "手搖",
            "夜市", "美食", "食品", "麵", "飯", "鍋",
            # 馬來西亞連鎖/外送
            "grab food", "grabfood", "food panda",
            "jollibee",
            "oldtown", "old town white coffee",
            "secret recipe",
            "marrybrown",
            "pizza hut", "domino",
            "nando", "nandos",
            "chatime", "tealive",
            "papa john",
            "subway",
            "a&w",
            "sugarbun",
            "ramly",
            "mamak", "nasi lemak", "roti canai",
            "restoran", "kedai makan",
            "warung",
            # 超市熟食
            "jaya grocer", "village grocer",
        ],
    },

    # ── 交通 ──────────────────────────────────────────────────────────────
    {
        "category": "交通",
        "keywords": [
            # 台灣
            "捷運", "mrt", "台鐵", "高鐵", "thsr", "客運", "公車", "巴士",
            "uber", "taxi", "計程車", "叫車", "youbike", "ubike",
            "加油", "中油", "台塑", "shell", "加油站",
            "停車", "parking", "停車費", "easycard", "悠遊卡",
            "機票", "航空", "airline", "飛機", "長榮", "華航", "eva air",
            "高速公路", "etag", "通行費",
            # 馬來西亞交通
            "grab car", "grabcar", "grab ride",
            "rapidkl", "rapidpenang", "rapid kl",
            "lrt", "mrt", "monorail", "ktm", "ets",
            "plus highway", "plus tol", "lebuhraya",
            "touch n go", "tng", "touch & go",
            "petronas", "petron", "bhp petrol", "bhpetrol", "caltex",
            "shell petrol",
            "airasia", "air asia", "malaysia airlines", "mas", "malindo",
            "firefly",
            "jetstar",
            "bas", "teksi",
            "e-hailing",
            "mycargo", "pos laju",
        ],
    },

    # ── 購物 ──────────────────────────────────────────────────────────────
    {
        "category": "購物",
        "keywords": [
            # 台灣
            "momo", "蝦皮", "shopee", "pchome", "yahoo購物", "博客來",
            "amazon", "淘寶", "taobao",
            "costco", "家樂福", "carrefour", "全聯", "大潤發", "愛買",
            "百貨", "sogo", "新光", "遠百", "微風", "101",
            "ikea", "特力屋", "b&q",
            "服飾", "衣服", "鞋子", "包包", "飾品",
            "3c", "電器", "手機", "電腦", "apple", "samsung",
            "書局", "誠品", "金石堂",
            # 馬來西亞電商/超市
            "lazada",
            "zalora",
            "tesco", "aeon", "aeon big", "giant", "99 speedmart",
            "mydin", "econsave", "lotus",
            "cold storage",
            "bens independent grocer",
            "hero market",
            "parkson", "robinsons",
            "isetan",
            "suria klcc", "mid valley", "sunway pyramid",
            "pavilion",
            "uniqlo", "h&m", "zara",
            "watson", "guardian pharmacy",
            "mr diy",
            "caring pharmacy",
            "popular bookstore",
            "times bookstore",
        ],
    },

    # ── 娛樂 ──────────────────────────────────────────────────────────────
    {
        "category": "娛樂",
        "keywords": [
            # 通用訂閱
            "netflix", "disney+", "disney plus", "hbo", "hbo go", "hbo max",
            "spotify", "kkbox", "apple music", "youtube premium",
            "twitch", "steam", "xbox", "playstation", "switch",
            "game", "遊戲",
            # 台灣
            "friday", "line tv", "hami video",
            "電影", "影院", "戲院", "威秀", "國賓", "cinemark",
            "ktv", "好樂迪", "錢櫃",
            "健身", "gym", "世界健身", "全國", "運動",
            "歌唱", "演唱會", "音樂會", "展覽", "博物館",
            # 馬來西亞
            "gsc cinema", "tgv cinema", "mmcineplexes", "lotus five star",
            "klpac", "istana budaya",
            "escape", "sunway lagoon", "legoland",
            "celebrity fitness", "true fitness", "anytime fitness",
            "badminton", "futsal",
            "karaoke",
        ],
    },

    # ── 醫療 ──────────────────────────────────────────────────────────────
    {
        "category": "醫療",
        "keywords": [
            # 台灣
            "醫院", "診所", "醫療", "門診", "掛號",
            "牙科", "牙醫", "眼科", "皮膚科", "骨科",
            "藥局", "藥房", "屈臣氏", "康是美",
            "保健", "維他命", "補品", "藥品",
            # 馬來西亞
            "hospital", "klinik", "clinic",
            "farmasi", "pharmacy",
            "guardian", "watson",
            "pantai hospital", "gleneagles", "sunway medical",
            "columbia asia", "imu",
            "kpj", "sime darby medical",
            "dentist", "dental",
            "optometrist", "optik",
            "specialist centre",
        ],
    },

    # ── 教育 ──────────────────────────────────────────────────────────────
    {
        "category": "教育",
        "keywords": [
            # 台灣
            "補習班", "課程", "學費", "學雜費", "學習",
            "udemy", "coursera", "hahow", "yotta",
            "書店", "教科書", "參考書", "文具",
            "家教", "tutor",
            # 馬來西亞
            "sekolah", "kolej", "universiti", "university",
            "yuran", "tuition", "buku teks",
            "sunway university", "taylor", "monash", "utar",
            "ipts", "ipta",
        ],
    },

    # ── 住宅 ──────────────────────────────────────────────────────────────
    {
        "category": "住宅",
        "keywords": [
            # 台灣
            "租金", "房租", "房貸",
            "台電", "電費", "水費", "台自來水", "瓦斯",
            "中華電信", "台灣大哥大", "遠傳", "網路費", "電話費",
            "管理費", "物業", "修繕", "裝潢",
            "保險", "壽險", "車險", "意外險",
            # 馬來西亞水電網
            "tenaga nasional", "tnb", "indah water", "syabas", "span",
            "maxis", "celcom", "digi", "u mobile", "yes 4g", "unifi",
            "time broadband", "astro",
            "sewa", "sewaan",
            "cukai tanah", "cukai pintu",
            "strata", "maintenance fee",
            "gas malaysia",
        ],
    },

    # ── 旅遊 ──────────────────────────────────────────────────────────────
    {
        "category": "旅遊",
        "keywords": [
            # 通用
            "飯店", "住宿", "hotel", "airbnb", "booking.com", "agoda",
            "旅遊", "觀光", "景點", "門票",
            "租車", "car rental",
            # 馬來西亞
            "resort", "chalet",
            "klcc", "petronas twin towers",
            "penang", "langkawi", "genting", "cameron highlands",
            "sabah", "sarawak",
            "visa", "passport",
        ],
    },

]


def categorize_transactions(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    批量分類交易記錄，就地修改 category 欄位

    Args:
        rows: 交易字典列表（需包含 description 與 is_income 欄位）

    Returns:
        同一列表（已修改 category）
    """
    for row in rows:
        if row.get("category") and row["category"] != "其他":
            continue  # 已有類別，跳過

        desc      = (row.get("description") or "").lower()
        is_income = bool(row.get("is_income", 0))

        row["category"] = _classify(desc, is_income)

    logger.debug("[Categorizer] 分類完成，共 %d 筆", len(rows))
    return rows


def _classify(desc_lower: str, is_income: bool) -> str:
    """對單筆交易進行分類"""
    for rule in _RULES:
        # 收入規則只適用於 is_income=True 的記錄
        if rule.get("is_income") and not is_income:
            continue
        for kw in rule["keywords"]:
            if kw.lower() in desc_lower:
                return rule["category"]

    return "收入" if is_income else "其他"


def get_category_for_description(description: str, is_income: bool = False) -> str:
    """對單一描述取得分類（供 API 呼叫）"""
    return _classify(description.lower(), is_income)
