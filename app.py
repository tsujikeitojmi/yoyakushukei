import streamlit as st
import pandas as pd
import uuid
import datetime
import unicodedata
import re
import json
import os
import difflib
from io import BytesIO
from supabase import create_client


_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(page_title="予約件数集計ツール", layout="wide")

# ---------------------------------------------------------------------------
# カスタム CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
/* ════════════════════════════════
   白背景・黒文字 統一テーマ
   ════════════════════════════════ */

/* 全体背景 */
.stApp, .main { background-color: #ffffff; }

/* ── テキスト入力 ── */
input, textarea,
[data-baseweb="input"] > div,
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
    border-color: #d8d8d8 !important;
}

/* ── セレクトボックス（閉じた状態） ── */
div[data-baseweb="select"] > div,
div[data-baseweb="select"] > div > div {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
    border-color: #d8d8d8 !important;
}
[data-baseweb="select"] span,
[data-baseweb="select"] div[class*="singleValue"],
[data-baseweb="select"] div[class*="placeholder"] {
    color: #1a1a1a !important;
}

/* ── ドロップダウンメニュー（開いた状態） ── */
[data-baseweb="menu"],
[data-baseweb="menu"] ul,
[data-baseweb="menu"] li,
[role="listbox"],
[role="option"] {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
}
[role="option"]:hover,
[data-baseweb="menu"] li:hover {
    background-color: #fff0f3 !important;
    color: #c94f6a !important;
}

/* 日付入力 */
[data-testid="stDateInput"] input,
[data-testid="stDateInput"] > div {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
}

/* フォーカス時 */
[data-baseweb="input"]:focus-within > div,
[data-baseweb="select"]:focus-within > div {
    border-color: #c94f6a !important;
    box-shadow: 0 0 0 2px rgba(201,79,106,.15) !important;
}

/* プレースホルダー */
::placeholder { color: #aaaaaa !important; opacity: 1; }

/* ── 見出し・一般テキスト ── */
h1, h2, h3, h4, h5, h6 {
    color: #1a1a1a !important;
}
label, p, li,
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p,
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3 {
    color: #1a1a1a !important;
}

/* ── メトリクスカード ── */
[data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #ebebeb;
    border-radius: 10px;
    padding: 14px 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
}
[data-testid="metric-container"] label {
    color: #888888 !important;
    font-size: 12px !important;
    letter-spacing: .04em;
}
[data-testid="stMetricValue"] {
    color: #1a1a1a !important;
    font-size: 24px !important;
    font-weight: 700 !important;
}
[data-testid="stMetricDelta"] { color: #888888 !important; font-size: 12px !important; }

/* ── タブ ── */
.stTabs [data-baseweb="tab-list"] {
    background: #ffffff;
    border-bottom: 2px solid #f0d0d8;
    gap: 2px;
}
.stTabs [data-baseweb="tab"] {
    background: #ffffff !important;
    color: #888888 !important;
    font-weight: 600;
    font-size: 14px;
    border-radius: 6px 6px 0 0;
    padding: 8px 20px;
}
.stTabs [aria-selected="true"] {
    background: #fff5f7 !important;
    color: #c94f6a !important;
    border-bottom: 2px solid #c94f6a !important;
}

/* ── ボタン（全種） ── */
button[kind="primary"],
button[kind="secondary"],
button[kind="tertiary"],
button[kind="minimal"],
.stButton > button,
.stFormSubmitButton > button,
[data-testid="stDownloadButton"] > button {
    background-color: #ffffff !important;
    color: #c94f6a !important;
    border: 1.5px solid #e8b0be !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    padding: 7px 18px !important;
    transition: background-color .15s, border-color .15s !important;
}
button[kind="primary"] p,
button[kind="secondary"] p,
.stButton > button p,
.stFormSubmitButton > button p,
[data-testid="stDownloadButton"] > button p {
    color: #c94f6a !important;
}
.stButton > button:hover,
.stFormSubmitButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    background-color: #fff0f3 !important;
    border-color: #c94f6a !important;
    color: #c94f6a !important;
}

/* ── フォーム枠 ── */
[data-testid="stForm"] {
    background-color: #fafafa !important;
    border: 1px solid #ebebeb !important;
    border-radius: 8px !important;
    padding: 16px 20px !important;
}

/* ── ファイルアップローダー ── */
[data-testid="stFileUploader"],
[data-testid="stFileUploaderDropzone"] {
    background-color: #fafafa !important;
    border: 1.5px dashed #d8d8d8 !important;
    border-radius: 8px !important;
    color: #1a1a1a !important;
}

/* ── データフレーム ── */
[data-testid="stDataFrame"] > div {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    overflow: hidden;
    background-color: #ffffff !important;
}
[data-testid="stDataFrame"] [data-testid="glideDataEditor"],
[data-testid="stDataFrame"] canvas {
    background-color: #ffffff !important;
}

/* ── 区切り線 ── */
hr { border-color: #ebebeb !important; margin: 16px 0 !important; }

/* ── キャプション ── */
small, [data-testid="stCaptionContainer"] p { color: #999999 !important; }

/* ── アラート ── */
[data-testid="stAlert"] { border-radius: 6px !important; }
[data-testid="stAlert"] p { color: inherit !important; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# タイトル
# ---------------------------------------------------------------------------
st.markdown("## 予約件数集計ツール")
st.divider()

# ---------------------------------------------------------------------------
# 定数・パス
# ---------------------------------------------------------------------------
COL_NAME     = "来院者名"
COL_START    = "予約開始日時"
COL_CANCEL   = "予約キャンセル日時"
COL_TYPE     = "初診・再診"
COL_CATEGORY = "施術カテゴリー"
COL_TREAT    = "施術名"
COL_MENU     = "メニュー導線"
COL_BIKO     = "備考"
COL_MENU_CONTENT = "メニュー内容"
COL_RESERVATION_ID = "予約ID"

REQUIRED_COLS = [COL_NAME, COL_START, COL_CANCEL, COL_TYPE]
CLINICS = ["心斎橋院", "新宿院", "福岡院"]

def _menu_price_file(clinic: str) -> str:
    return os.path.join(_DIR, f"menu_price_cache_{clinic}.csv")

def _menu_price_name_file(clinic: str) -> str:
    return os.path.join(_DIR, f"menu_price_cache_{clinic}_name.txt")

_MASTER_XLSX = os.path.join(os.path.expanduser("~"), "Downloads", "媒体掲載メニュー一覧.xlsx")

CLINIC_SHEET_MAP = {
    "心斎橋院": "心斎橋",
    "新宿院":   "新宿",
    "福岡院":   "福岡",
    "名古屋院": "名古屋",
}


def load_master_xlsx(clinic: str):
    """媒体掲載メニュー一覧.xlsxから院別シートを読み込む。(df, error_msg) を返す。"""
    sheet = CLINIC_SHEET_MAP.get(clinic)
    if sheet is None:
        return None, f"「{clinic}」のシートマッピングが未登録です。"
    if not os.path.exists(_MASTER_XLSX):
        return None, f"ファイルが見つかりません：{_MASTER_XLSX}"
    try:
        df = pd.read_excel(_MASTER_XLSX, sheet_name=sheet, dtype=str)
        if "メニュー名" not in df.columns:
            return None, "「メニュー名」列が見つかりません。"
        if "参考料金" not in df.columns:
            return None, "「参考料金」列が見つかりません。"
        df = df[["メニュー名", "参考料金"]].copy()
        df = df[df["メニュー名"].notna() & (df["メニュー名"].astype(str).str.strip() != "")]
        return df.reset_index(drop=True), None
    except Exception as e:
        return None, f"読み込みエラー：{e}"


def build_price_dict_from_master(master_df: pd.DataFrame) -> dict:
    """メニュー情報CSV → {正規化キー: {"name": ..., "price": ...}} の辞書"""
    d = {}
    for _, row in master_df.iterrows():
        name_raw = row.get("メニュー名※必須", "")
        if pd.isna(name_raw) or str(name_raw).strip() == "":
            continue
        name  = str(name_raw).strip()
        price = parse_price(row.get("金額", None))
        key1  = normalize_menu(name)
        key2  = normalize_menu(strip_menu_prefix(name))
        entry = {"name": name, "price": price}
        if key1:
            if key1 not in d or d[key1]["price"] is None:
                d[key1] = entry
        if key2 and key2 != key1:
            if key2 not in d or d[key2]["price"] is None:
                d[key2] = entry
    return d


# ---------------------------------------------------------------------------
# 金額集計ロジック
# ---------------------------------------------------------------------------

def parse_price(price_str):
    """
    金額文字列をパースして整数を返す。
    例: "19 | 900円" → 19900
        "9800円" → 9800
        "24 | 800円（通常価格39 | 800円）" → 24800
        "19 | 800円~" → 19800
    """
    if pd.isna(price_str) or str(price_str).strip() == "":
        return None
    s = str(price_str)
    # → / ⇒ がある場合は最後の部分（キャンペーン価格等）を使う
    for arrow in ('→', '⇒'):
        if arrow in s:
            s = s.split(arrow)[-1]
            break
    def _extract(digits_str):
        cleaned = re.sub(r'[\s|｜,，]+', '', digits_str)
        return int(cleaned) if cleaned.isdigit() else None

    # キャンペーン価格が（）内にある場合はそちらを優先
    _campaign_kws = ['キャンペーン', 'CP', 'セール', '割引', '特価']
    for paren_content in re.findall(r'[（(]([^）)]*)[）)]', s):
        if any(kw in paren_content for kw in _campaign_kws):
            m = re.search(r'(\d[\d\s|｜,，]*)\s*円', paren_content)
            if m:
                result = _extract(m.group(1))
                if result is not None:
                    return result

    # （...）を除去
    s = re.sub(r'[（(][^）)]*[）)]', '', s)

    # パターン1: 「円」の直前の数字（| 区切り・カンマ含む）
    m = re.search(r'(\d[\d\s|｜,，]*)\s*円', s)
    if m:
        result = _extract(m.group(1))
        if result is not None:
            return result

    # パターン2: ￥ または ¥ の直後の数字（| 区切り・カンマ含む）
    m = re.search(r'[￥¥]\s*(\d[\d\s|｜,，]*)', s)
    if m:
        result = _extract(m.group(1))
        if result is not None:
            return result

    return None


def normalize_menu(text) -> str:
    """メニュー名照合用の正規化"""
    if pd.isna(text):
        return ""
    t = unicodedata.normalize("NFKC", str(text)).strip()
    t = re.sub(r'\s+', '', t)
    return t.lower()


def strip_menu_prefix(text: str) -> str:
    """【...】〈...〉《...》〔...〕や★〇などの装飾を除去"""
    t = re.sub(r'【[^】]*】', '', text)
    t = re.sub(r'〔[^〕]*〕', '', t)
    t = re.sub(r'[〈《][^〉》]*[〉》]', '', t)
    # 二重括弧等で残った先頭の閉じ括弧を除去
    t = re.sub(r'^[】〕〉》\]）)]+', '', t)
    t = re.sub(r'[★☆〇●◎◆◇■□▲△※♦♣♠♥✨]', '', t)
    # ※以降の補足文を除去
    t = re.sub(r'※.*$', '', t)
    return t.strip()




# ---------------------------------------------------------------------------
# サイト料金表スクレイピング（最終フォールバック用）
# ---------------------------------------------------------------------------
_SITE_PRICE_URL   = "https://dailyskinclinic.jp/price"
_SITE_CLINIC_ORDER = ["新宿院", "心斎橋院", "名古屋院", "福岡院"]


@st.cache_data(ttl=86400, show_spinner=False)
def scrape_site_prices() -> dict:
    """dailyskinclinic.jpの料金ページを院別にスクレイピングして
    {院名: {正規化メニュー名: 価格}} を返す。キャンペーン価格を優先。"""
    try:
        import urllib.request
        from bs4 import BeautifulSoup

        req = urllib.request.Request(_SITE_PRICE_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as res:
            html = res.read().decode("utf-8", errors="ignore")

        soup = BeautifulSoup(html, "html.parser")
        sections = soup.find_all(class_="price-main-contents")
        result = {}

        for i, section in enumerate(sections):
            if i >= len(_SITE_CLINIC_ORDER):
                break
            clinic = _SITE_CLINIC_ORDER[i]
            tables = section.find_all("table", class_="price-main-contents-menu-price-table")
            entries = {}  # {norm_key: (raw_name, price)}

            for table in tables:
                current_name = None
                current_price = None
                current_pri   = 0

                def _try_add():
                    if current_name and current_price:
                        key = normalize_menu(current_name)
                        if key and key not in entries:
                            entries[key] = (current_name, current_price)

                def _price_priority(price_type_text, price_val):
                    if "キャンペーン" in price_type_text:
                        return price_val, 3
                    if "初回" in price_type_text:
                        return price_val, 2
                    if "通常" in price_type_text:
                        return price_val, 1
                    return price_val, 0

                for row in table.find_all("tr"):
                    cells = row.find_all("td")
                    if not cells:
                        continue

                    if cells[0].get("rowspan"):
                        _try_add()
                        td = cells[0]
                        for font in td.find_all("font"):
                            font.decompose()
                        current_name  = td.get_text(separator=" ", strip=True)
                        current_price = None
                        current_pri   = 0
                        if len(cells) >= 3:
                            p = parse_price(cells[2].get_text(strip=True))
                            if p and p > 0:
                                current_price, current_pri = _price_priority(cells[1].get_text(strip=True), p)
                    elif len(cells) == 2 and current_name:
                        p = parse_price(cells[1].get_text(strip=True))
                        if p and p > 0:
                            val, pri = _price_priority(cells[0].get_text(strip=True), p)
                            if pri > current_pri:
                                current_price, current_pri = val, pri

                _try_add()
            result[clinic] = entries

        return result
    except Exception:
        return {}


def lookup_site_price(menu_raw: str, clinic_name: str, site_prices: dict):
    """サイト料金表から照合して (site_name, price) or None を返す。
    ①正規化後の部分一致 → ②単語セット一致（順番違い・スペース差異対応）の順で試みる。"""
    entries = site_prices.get(clinic_name, {})
    if not entries:
        return None
    _TYPO_MAP = {"blow": "brow", "bloe": "brow"}
    # 連結英単語を分解（例: powderbrow → powder brow）
    _COMPOUND_SPLIT = [
        (r'(?i)(powder)(brow)', r'\1 \2'),
        (r'(?i)(natural)(brow)', r'\1 \2'),
        (r'(?i)(daily)(brow)',   r'\1 \2'),
    ]
    stripped = strip_menu_prefix(menu_raw)
    for wrong, correct in _TYPO_MAP.items():
        stripped = re.sub(wrong, correct, stripped, flags=re.IGNORECASE)
    for pattern, repl in _COMPOUND_SPLIT:
        stripped = re.sub(pattern, repl, stripped)
    norm = normalize_menu(stripped)
    if not norm or len(norm) < 4:
        return None

    def _word_set(raw_text):
        parts = re.split(r'[\s　・|｜/／,、。\-]+', str(raw_text))
        return set(normalize_menu(p) for p in parts if len(normalize_menu(p)) >= 3)

    query_words = _word_set(stripped)

    best_price, best_raw, best_score = None, None, 0
    for site_norm_key, (site_raw_name, price) in entries.items():
        # ①正規化後の部分一致
        if norm in site_norm_key:
            score = len(norm) * 2
        elif site_norm_key in norm:
            score = len(site_norm_key) * 2
        else:
            # ②単語セット一致（空白除去前の名前で分割）
            site_words = _word_set(site_raw_name)
            if not site_words or not query_words:
                continue
            overlap = query_words & site_words
            if not overlap:
                continue
            score = sum(len(w) for w in overlap)
            if score < 4:
                continue
        if score > best_score:
            best_score, best_price, best_raw = score, price, site_raw_name
    if best_price and best_score >= 4:
        return best_raw, best_price
    return None


def build_price_dict(menu_csv_df: pd.DataFrame) -> dict:
    """
    メニュー情報CSV → {正規化キー: {"name": ..., "price": ...}} の辞書
    列: メニュー名※必須, 金額
    """
    d = {}
    col_name  = "メニュー名※必須"
    col_price = "金額"
    for _, row in menu_csv_df.iterrows():
        name_raw = row.get(col_name, "")
        if pd.isna(name_raw) or str(name_raw).strip() == "":
            continue
        name  = str(name_raw).strip()
        price = parse_price(row.get(col_price, None))
        # 正規化キー
        key1 = normalize_menu(name)
        key2 = normalize_menu(strip_menu_prefix(name))
        entry = {"name": name, "price": price}
        if key1:
            # 既存エントリに金額がない場合は上書き
            if key1 not in d or d[key1]["price"] is None:
                d[key1] = entry
        if key2 and key2 != key1:
            # 既存エントリに金額がない場合は上書き
            if key2 not in d or d[key2]["price"] is None:
                d[key2] = entry
    return d



def build_treatment_dict_from_csv(menu_csv_df: pd.DataFrame) -> dict:
    """
    「メニューに含める施術※必須」列を | で分割し、
    {正規化施術名: {"name": メニュー名, "price": 金額}} の辞書を返す。
    同じ施術名が複数メニューにある場合は最安値を採用。
    """
    col_name      = "メニュー名※必須"
    col_price     = "金額"
    col_treatment = "メニューに含める施術※必須"
    d = {}
    if col_treatment not in menu_csv_df.columns:
        return d
    for _, row in menu_csv_df.iterrows():
        treat_raw = row.get(col_treatment, "")
        if pd.isna(treat_raw) or str(treat_raw).strip() == "":
            continue
        name_raw = row.get(col_name, "")
        if pd.isna(name_raw) or str(name_raw).strip() == "":
            continue
        price = parse_price(row.get(col_price, None))
        if price is None or price <= 0:
            continue
        entry = {"name": str(name_raw).strip(), "price": price}
        for part in str(treat_raw).split("|"):
            part = part.strip()
            if not part:
                continue
            if any(kw in part for kw in ['カウンセリング', '洗顔', '麻酔']):
                continue
            for key in [normalize_menu(part), normalize_menu(strip_menu_prefix(part))]:
                if not key:
                    continue
                if key not in d or price < d[key]["price"]:
                    d[key] = entry
    return d


FIXED_PRICE_KEYWORDS = [
    ("イソトレチノイン",         9800),
    ("マンジャロ",              15900),
    ("ストレッチマーク2回コース",  70000),
    ("ストレッチマーク1回",       39800),
    ("傷跡2×2cm以内 1回",       9800),
    ("傷跡10×10cm以内 1回",     23800),
    ("Skn52 フリー",             23800),
    ("10cm×10cm以内 1回",       23800),
    ("10cm×10cm以内 2回",       42800),
    ("dazzybrow 2回",           92400),
    ("dazzyblow 2回",           92400),
    ("dazzybrow 1回",           55000),
    ("dazzyblow 1回",           55000),
    ("アイライン上",              46200),
    ("daily brow",              92400),
    ("daily lip",               92400),
    ("daily blow",              92400),
    ("natural brow",            69300),
    ("natural blow",            69300),
    ("powder brow",             69300),
    ("powder blow",             69300),
    ("ショートスレッド20本",      39800),
    ("最強リフトアップセット",    59800),
]

# 複数キーワードをすべて含む場合に固定金額を使用（AND条件）
FIXED_PRICE_KEYWORDS_AND = [
    (["女性", "タイムチャージ"], 19800),
    (["男性", "タイムチャージ"], 29800),
]

FIXED_PRICE_EXACT = {
    "ハイドラ": 8800,
}

def match_price(menu_content: str, price_dict: dict, exact_only: bool = False) -> dict:
    """
    メニュー内容1件を価格辞書と照合。
    exact_only=True の場合はステップ1（正規化完全一致）のみ実行。
    returns: {"matched_name", "price", "status", "raw"}
      status: "一致" / "未一致"
    """
    raw = str(menu_content).strip()

    # 固定価格（完全一致・正規化一致）
    _norm_raw_exact = normalize_menu(raw)
    for _ek, _ev in FIXED_PRICE_EXACT.items():
        if raw == _ek or _norm_raw_exact == normalize_menu(_ek):
            return {"raw": raw, "matched_name": _ek, "price": _ev, "status": "一致"}

    # 固定価格キーワード AND条件
    _raw_lower = raw.lower()
    for kws, fixed_price in FIXED_PRICE_KEYWORDS_AND:
        if all(kw in _raw_lower for kw in kws):
            return {"raw": raw, "matched_name": raw, "price": fixed_price, "status": "一致"}

    # 固定価格キーワード（量によって金額が変わるため固定値を使用）
    _norm_raw = normalize_menu(raw)
    for kw, fixed_price in FIXED_PRICE_KEYWORDS:
        if kw in raw or normalize_menu(kw) in _norm_raw:
            return {"raw": raw, "matched_name": raw, "price": fixed_price, "status": "一致"}

    def _try(key):
        return price_dict.get(key)

    # ステップ1: 正規化完全一致（金額がある場合のみ採用）
    k = normalize_menu(raw)
    if e := _try(k):
        if e["price"] is not None:
            return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}

    if exact_only:
        return {"raw": raw, "matched_name": None, "price": None, "status": "未一致"}

    # ステップ2: 装飾除去後
    k2 = normalize_menu(strip_menu_prefix(raw))
    if k2 and (e := _try(k2)):
        if e["price"] is not None:
            return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}

    # ステップ3: カウンセリング+ / カウンセリング＋ 前置きを除去して照合
    for prefix in ('カウンセリング+', 'カウンセリング＋'):
        if raw.startswith(prefix):
            rest = raw[len(prefix):].strip()
            k3 = normalize_menu(rest)
            if k3 and (e := _try(k3)):
                if e["price"] is not None:
                    return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
            k3b = normalize_menu(strip_menu_prefix(rest))
            if k3b and (e := _try(k3b)):
                if e["price"] is not None:
                    return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
            break

    # ステップ4: or / or で複数選択肢がある場合、最初の選択肢で照合
    for sep in ('or', 'or'):
        if sep in normalize_menu(raw):
            first = normalize_menu(raw).split(sep)[0].strip()
            if first and (e := _try(first)):
                if e["price"] is not None:
                    return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
            # 装飾除去後でも試す
            raw_stripped = strip_menu_prefix(raw)
            first2 = normalize_menu(raw_stripped).split(sep)[0].strip()
            if first2 and (e := _try(first2)):
                if e["price"] is not None:
                    return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
            # 最後の選択肢の後続語（例: "脇 100単位" の "100単位"）を先頭選択肢に付加して照合
            # 例: "ボツラックス肩or脇 100単位" → "ボツラックス肩 100単位"
            raw_parts = re.split(r'or|or', raw_stripped)
            if len(raw_parts) >= 2:
                raw_first_part = raw_parts[0].strip()
                raw_last_words = raw_parts[-1].strip().split()
                if len(raw_last_words) > 1:
                    trailing = " ".join(raw_last_words[1:])
                    candidate = normalize_menu(raw_first_part + " " + trailing)
                    if candidate and (e := _try(candidate)):
                        if e["price"] is not None:
                            return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
            break

    # ステップ5: （...）の処理
    _base = strip_menu_prefix(raw)
    # 5a: 括弧を外して中身を残す（デンシティ（全顔300ショット）→デンシティ全顔300ショット）
    k5a = normalize_menu(re.sub(r'[（(]([^）)]*)[）)]', r'\1', _base).strip())
    if k5a and (e := _try(k5a)):
        if e["price"] is not None:
            return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
    # 5b: 括弧ごと削除（ポテンツァ（マックーム）→ポテンツァマックーム用）
    k5b = normalize_menu(re.sub(r'[（(][^）)]*[）)]', '', _base).strip())
    if k5b and (e := _try(k5b)):
        if e["price"] is not None:
            return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}
    # 5c: 括弧内から「数字+ショット/shot」だけ抽出して結合（デンシティ（全顔300ショット）→デンシティ300ショット）
    _paren_m = re.search(r'[（(][^）)]*[）)]', _base)
    if _paren_m:
        _shot = re.search(r'(\d+\s*[sｓ][hｈ][oｏ][tｔ]|\d+\s*ショット)', _paren_m.group(), re.IGNORECASE)
        if _shot:
            _base_no_paren = re.sub(r'[（(][^）)]*[）)]', '', _base).strip()
            k5c = normalize_menu(_base_no_paren + _shot.group().strip())
            if k5c and (e := _try(k5c)):
                if e["price"] is not None:
                    return {"raw": raw, "matched_name": e["name"], "price": e["price"], "status": "一致"}

    # ステップ5d: メニュー名に金額が直書きされている場合（例: ダーマペン(9 | 800円単体時））
    _m = re.search(r'(\d[\d\s|｜]*)\s*円', raw)
    if _m:
        _cleaned = re.sub(r'[\s|｜,，]+', '', _m.group(1))
        if _cleaned.isdigit():
            return {"raw": raw, "matched_name": raw, "price": int(_cleaned), "status": "内容より取得"}
    _m = re.search(r'[￥¥]\s*(\d[\d\s|｜]*)', raw)
    if _m:
        _cleaned = re.sub(r'[\s|｜,，]+', '', _m.group(1))
        if _cleaned.isdigit():
            return {"raw": raw, "matched_name": raw, "price": int(_cleaned), "status": "内容より取得"}

    # ステップ6: 部分一致（正規化マスタ名が正規化予約名に含まれる）
    k_raw = normalize_menu(raw)
    k_stripped = normalize_menu(strip_menu_prefix(raw))
    best_len = 0
    best_entry = None
    for key, entry in price_dict.items():
        if entry["price"] is None or not key:
            continue
        if key in k_raw or key in k_stripped:
            if len(key) > best_len:
                best_len = len(key)
                best_entry = entry
    if best_entry is not None:
        return {"raw": raw, "matched_name": best_entry["name"], "price": best_entry["price"],
                "status": "部分一致"}

    # ステップ7: あいまい一致（fuzzy）
    k_target = k_stripped or k_raw
    if k_target:
        best_score = 0.0
        best_fuzz = None
        for key, entry in price_dict.items():
            if entry["price"] is None:
                continue
            score = difflib.SequenceMatcher(None, k_target, key).ratio()
            if score > best_score:
                best_score = score
                best_fuzz = entry
        if best_score >= 0.80 and best_fuzz is not None:
            return {"raw": raw, "matched_name": best_fuzz["name"], "price": best_fuzz["price"],
                    "status": f"あいまい一致({int(best_score*100)}%)"}

    return {"raw": raw, "matched_name": None, "price": None, "status": "未一致"}


def aggregate_prices(
    res_df: pd.DataFrame,
    price_dict: dict,
    date_from: datetime.date,
    date_to: datetime.date,
    treat_csv_dict: dict = None,
    clinic_name: str = "",
) -> tuple:
    """
    予約CSVを処理して金額集計の明細DFを返す。
    returns: (detail_df, unmatched_df, empty_df)
    """
    df = res_df.copy()
    df["_start_dt"]   = pd.to_datetime(df[COL_START], errors="coerce")
    df["_start_date"] = df["_start_dt"].dt.date

    # 来院者名が空白の行を除外
    if COL_NAME in df.columns:
        df = df[df[COL_NAME].notna() & (df[COL_NAME].astype(str).str.strip() != "")]

    # キャンセル除外（無条件）
    if COL_CANCEL in df.columns:
        df = df[df[COL_CANCEL].isna() | (df[COL_CANCEL].astype(str).str.strip() == "")]

    # 日付フィルター
    df = df[(df["_start_date"] >= date_from) & (df["_start_date"] <= date_to)]

    # 補助項目（カウンセリング・麻酔・洗顔・モニター）を含む行を除外
    if COL_TREAT in df.columns:
        df = df[~df[COL_TREAT].apply(
            lambda s: any(kw in str(s) for kw in ['カウンセリング', '麻酔', '洗顔', 'モニター撮影', 'モニター写真', '物販'])
            if not pd.isna(s) else False
        )]

    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    records   = []
    unmatched = []
    empty_cnt_inner = 0

    for res_id, group in df.groupby(COL_RESERVATION_ID, sort=False):
        first_row   = group.sort_values(COL_START).iloc[0]
        start_dt    = first_row["_start_dt"]
        visitor     = first_row.get(COL_NAME, "")

        # メニュー内容を収集し、|で分割して個別メニューに展開
        raw_cells = group[COL_MENU_CONTENT].dropna().astype(str).tolist()
        raw_cells = [c for c in raw_cells if c.strip()]  # 空文字も除外
        raw_cells = list(dict.fromkeys(raw_cells))  # セル単位で重複除去

        # 備考から金額を抽出するヘルパー
        def _price_from_biko():
            if COL_BIKO not in group.columns:
                return None
            for biko_val in group[COL_BIKO].dropna().astype(str).tolist():
                p = parse_price(biko_val)
                if p is not None and p > 0:
                    return p
            return None

        # メニュー内容が空欄の場合 → 施術名マッピングで照合
        if not raw_cells:
            empty_cnt_inner += 1
            total_price    = 0
            matched_names  = []
            matched_prices = []
            treat_cells = group[COL_TREAT].dropna().astype(str).tolist()
            treat_cells = list(dict.fromkeys([t.strip() for t in treat_cells if t.strip()]))
            treat_cells = [t for t in treat_cells
                           if not any(kw in t for kw in ['カウンセリング', '麻酔', '洗顔'])]
            for tc in treat_cells:
                # 固定価格（完全一致）
                _norm_tc_exact = normalize_menu(tc)
                _tc_exact_price = next((ev for ek, ev in FIXED_PRICE_EXACT.items() if tc == ek or _norm_tc_exact == normalize_menu(ek)), None)
                if _tc_exact_price is not None:
                    matched_names.append(tc)
                    matched_prices.append(_tc_exact_price)
                    total_price += _tc_exact_price
                    continue
                # 固定価格キーワードを施術名にも適用（AND条件）
                for kws, fixed_price in FIXED_PRICE_KEYWORDS_AND:
                    if all(kw in tc.lower() for kw in kws):
                        matched_names.append(tc)
                        matched_prices.append(fixed_price)
                        total_price += fixed_price
                        break
                else:
                    _norm_tc = normalize_menu(tc)
                    for kw, fixed_price in FIXED_PRICE_KEYWORDS:
                        if kw in tc or normalize_menu(kw) in _norm_tc:
                            matched_names.append(tc)
                            matched_prices.append(fixed_price)
                            total_price += fixed_price
                            break
                    else:
                        if treat_csv_dict:
                            k_tc = normalize_menu(tc)
                            k_tc_s = normalize_menu(strip_menu_prefix(tc))
                            entry = treat_csv_dict.get(k_tc) or treat_csv_dict.get(k_tc_s)
                            if entry is None:
                                # 部分一致（treat_dictのキーが施術名に含まれる）
                                best_len, entry = 0, None
                                for key, e in treat_csv_dict.items():
                                    if key and (key in k_tc or key in k_tc_s):
                                        if len(key) > best_len:
                                            best_len, entry = len(key), e
                            if entry is not None and entry["price"] > 0:
                                matched_names.append(entry["name"])
                                matched_prices.append(entry["price"])
                                total_price += entry["price"]
                                continue
                        # treat_dictで取れなかった場合 → 施術名をメニュー名※必須と照合
                        if price_dict:
                            result = match_price(tc, price_dict)
                            if result["status"] != "未一致" and result["price"]:
                                matched_names.append(result["matched_name"])
                                matched_prices.append(result["price"])
                                total_price += result["price"]
            # 最終フォールバック: 備考 → サイト料金表
            if total_price == 0:
                biko_price = _price_from_biko()
                if biko_price:
                    total_price = biko_price
                    matched_names.append("備考より取得")
                    matched_prices.append(biko_price)
            records.append({
                "予約ID":         res_id,
                "予約開始日時":   start_dt,
                "予約日":         start_dt.date() if pd.notna(start_dt) else None,
                "来院者名":       visitor,
                "元メニュー内容": "",
                "施術名":         " / ".join(treat_cells),
                "一致メニュー名": " / ".join(matched_names),
                "採用金額一覧":   " / ".join([str(p) for p in matched_prices]),
                "照合元":         "施術名",
                "予約合計金額":   total_price,
            })
            continue

        menu_items = []
        for cell in raw_cells:
            cell_s = cell.strip()
            # セル自体が価格文字列の場合（例: "12 | 000円"）は | 分割せずそのまま保持
            if ('円' in cell_s or '¥' in cell_s or '￥' in cell_s) and parse_price(cell_s) is not None:
                menu_items.append(cell_s)
            else:
                parts = [p.strip() for p in cell_s.split('|') if p.strip()]
                menu_items.extend(parts if parts else [cell_s])
        # 展開後も重複除去
        menu_items = list(dict.fromkeys(menu_items))
        # 補助項目（カウンセリング・洗顔）をメニュー名でも除外
        # ただし price_dict に一致する複合メニュー名（例: カウンセリング+ハイコックス...）は除外しない
        menu_items = [m for m in menu_items
                      if match_price(m, price_dict, exact_only=True)["status"] == "一致"
                      or not any(kw in m for kw in ['カウンセリング', '洗顔'])]

        total_price   = 0
        matched_names = []
        matched_prices = []
        match_statuses = []
        has_unmatched  = False

        # フェーズ1: 完全一致のみ
        phase1_unmatched = []
        for mc in menu_items:
            result = match_price(mc, price_dict, exact_only=True)
            if result["status"] == "一致":
                matched_names.append(result["matched_name"])
                p = result["price"] if result["price"] is not None else 0
                matched_prices.append(p)
                match_statuses.append("一致")
                total_price += p
            else:
                phase1_unmatched.append(mc)

        # フェーズ3: フェーズ1で未一致のアイテムをフルマッチング
        for mc in phase1_unmatched:
            result = match_price(mc, price_dict)
            if result["status"] != "未一致":
                matched_names.append(result["matched_name"])
                p = result["price"] if result["price"] is not None else 0
                matched_prices.append(p)
                match_statuses.append(result["status"])
                total_price += p
            elif treat_csv_dict:
                # フェーズ3b: メニュー名※必須で取れなければ「メニューに含める施術※必須」でフォールバック
                mc_stripped = strip_menu_prefix(mc)
                k1 = normalize_menu(mc)
                k2 = normalize_menu(mc_stripped)
                entry = treat_csv_dict.get(k1) or treat_csv_dict.get(k2)
                if entry is None:
                    best_len, entry = 0, None
                    for key, e in treat_csv_dict.items():
                        if key and (key in k1 or key in k2 or k1 in key or k2 in key):
                            if len(key) > best_len:
                                best_len, entry = len(key), e
                if entry and entry["price"] > 0:
                    matched_names.append(entry["name"])
                    matched_prices.append(entry["price"])
                    match_statuses.append("施術CSV")
                    total_price += entry["price"]
                else:
                    has_unmatched = True
                    unmatched.append({
                        "予約ID":         res_id,
                        "予約日":         start_dt.date() if pd.notna(start_dt) else None,
                        "来院者名":       visitor,
                        "未一致メニュー": mc,
                    })
            else:
                has_unmatched = True
                unmatched.append({
                        "予約ID":         res_id,
                        "予約日":         start_dt.date() if pd.notna(start_dt) else None,
                        "来院者名":       visitor,
                        "未一致メニュー": mc,
                    })

        # 最終フォールバック: 備考 → サイト料金表
        if total_price == 0:
            biko_price = _price_from_biko()
            if biko_price:
                total_price = biko_price
                matched_names.append("備考より取得")
                matched_prices.append(biko_price)
                match_statuses.append("備考")

        records.append({
            "予約ID":         res_id,
            "予約開始日時":   start_dt,
            "予約日":         start_dt.date() if pd.notna(start_dt) else None,
            "来院者名":       visitor,
            "元メニュー内容": " / ".join(raw_cells),
            "一致メニュー名": " / ".join(matched_names),
            "採用金額一覧":   " / ".join([str(p) for p in matched_prices]),
            "照合方法":       " / ".join(match_statuses),
            "照合元":         "メニュー内容",
            "予約合計金額":   total_price,
        })

    detail_df    = pd.DataFrame(records)
    unmatched_df = pd.DataFrame(unmatched)
    return detail_df, unmatched_df, empty_cnt_inner


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Supabase クライアント
# ---------------------------------------------------------------------------
@st.cache_resource
def get_supabase():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_ANON_KEY"],
    )


# ---------------------------------------------------------------------------
# 条件マスタ CRUD（Supabase）
# ---------------------------------------------------------------------------
def load_all_conditions() -> dict:
    res = get_supabase().table("conditions").select("*").order("sort_order").execute()
    all_conds = {c: [] for c in CLINICS}
    for row in res.data:
        clinic = row["clinic"]
        if clinic in all_conds:
            all_conds[clinic].append({
                "id":           row["id"],
                "name":         row["name"],
                "category":     row.get("category") or "",
                "treatment":    row.get("treatment") or "",
                "sort_order":   row.get("sort_order") or 0,
                "campaign_end": row.get("campaign_end"),
            })
    return all_conds


def add_condition(clinic: str, cond: dict):
    get_supabase().table("conditions").insert({
        "id":           cond["id"],
        "clinic":       clinic,
        "name":         cond["name"],
        "category":     cond.get("category") or "",
        "treatment":    cond.get("treatment") or "",
        "sort_order":   cond.get("sort_order") or 0,
        "campaign_end": cond.get("campaign_end"),
    }).execute()


def update_condition(cond: dict):
    get_supabase().table("conditions").update({
        "name":         cond["name"],
        "category":     cond.get("category") or "",
        "treatment":    cond.get("treatment") or "",
        "campaign_end": cond.get("campaign_end"),
    }).eq("id", cond["id"]).execute()


def delete_condition(cond_id: str):
    get_supabase().table("conditions").delete().eq("id", cond_id).execute()


def delete_all_conditions(clinic: str):
    get_supabase().table("conditions").delete().eq("clinic", clinic).execute()


def load_custom_order() -> dict:
    try:
        res = get_supabase().table("settings").select("value").eq("key", "custom_order").execute()
        if res.data:
            return json.loads(res.data[0]["value"])
    except Exception:
        pass
    return {}


def save_custom_order(order: dict):
    try:
        get_supabase().table("settings").upsert({
            "key":   "custom_order",
            "value": json.dumps(order, ensure_ascii=False),
        }).execute()
    except Exception:
        pass



# ---------------------------------------------------------------------------
# セッションステート初期化
# ---------------------------------------------------------------------------
if "all_conditions" not in st.session_state:
    st.session_state["all_conditions"] = load_all_conditions()
if "edit_id" not in st.session_state:
    st.session_state["edit_id"] = None
if "bulk_delete_step" not in st.session_state:
    st.session_state["bulk_delete_step"] = {}  # { clinic_name: 0/1/2 }
if "selected_ids" not in st.session_state:
    st.session_state["selected_ids"] = {}  # { clinic_name: set of ids }
if "custom_cond_order" not in st.session_state:
    st.session_state["custom_cond_order"] = load_custom_order()
if "cond_not_found" not in st.session_state:
    st.session_state["cond_not_found"] = []


# ---------------------------------------------------------------------------
# フィルターパイプライン
# ---------------------------------------------------------------------------
def filter_cancelled(df: pd.DataFrame):
    mask = df[COL_CANCEL].isna() | (df[COL_CANCEL].astype(str).str.strip() == "")
    dropped = (~mask).sum()
    return df[mask].copy(), "キャンセル行を除外", dropped


def filter_duplicates(df: pd.DataFrame):
    before = len(df)
    df = df.drop_duplicates(subset=[COL_START, COL_NAME])
    dropped = before - len(df)
    return df.copy(), "重複行を正規化", dropped


def filter_no_name(df: pd.DataFrame):
    mask = df[COL_NAME].notna() & (df[COL_NAME].astype(str).str.strip() != "")
    dropped = (~mask).sum()
    return df[mask].copy(), "来院者名が空の行を除外", dropped


EXTRA_FILTERS = [filter_no_name]


def run_pipeline(df: pd.DataFrame):
    steps = [("元データ", len(df), None)]

    # キャンセル除外（施術別集計はここから使う）
    df, label, dropped = filter_cancelled(df)
    steps.append((label, len(df), dropped))
    after_cancel_df = df.copy()

    # 重複除外
    df, label, dropped = filter_duplicates(df)
    steps.append((label, len(df), dropped))

    for fn in EXTRA_FILTERS:
        df, label, dropped = fn(df)
        steps.append((label, len(df), dropped))

    steps.append(("最終件数", len(df), None))
    return df, after_cancel_df, steps


# ---------------------------------------------------------------------------
# テキスト正規化（全角→半角・スペース除去）
# ---------------------------------------------------------------------------
def normalize(text: str) -> str:
    """全角/半角を統一し、スペース（全角・半角）をすべて除去する"""
    t = unicodedata.normalize("NFKC", str(text))  # 全角英数記号→半角
    t = t.replace("\u3000", "").replace(" ", "")   # 全角スペース・半角スペース除去
    return t.lower()                                # 大文字小文字も統一


# ---------------------------------------------------------------------------
# 条件マッチング（部分一致・正規化済み）
# ---------------------------------------------------------------------------
def matches_condition(row: pd.Series, cond: dict) -> bool:
    cat_val   = normalize(row.get(COL_CATEGORY) or "")
    treat_val = normalize(row.get(COL_TREAT)    or "")
    cat_cond   = normalize(cond.get("category",  "") or "")
    treat_cond = normalize(cond.get("treatment", "") or "")
    cat_ok   = (not cat_cond)   or (cat_cond   in cat_val)
    treat_ok = (not treat_cond) or (treat_cond in treat_val)
    return cat_ok and treat_ok


# ---------------------------------------------------------------------------
# ヘルパー
# ---------------------------------------------------------------------------
def load_csv(file) -> pd.DataFrame:
    raw = file.read()
    for enc in ("utf-8-sig", "utf-8", "shift_jis", "cp932"):
        try:
            return pd.read_csv(BytesIO(raw), encoding=enc, dtype=str)
        except Exception:
            continue
    raise ValueError("CSVの読み込みに失敗しました。")


def classify_row(row: pd.Series) -> str:
    """
    判定順:
    1. 初診・再診 列に「初診」or「再診」が含まれる
    2. メニュー導線 に「LP用」「はじめてのDAILY」「初診」「再診」が含まれる
    3. メニュー内容 に「初診」or「再診」が含まれる
    4. どれにも該当しない → その他
    """
    def check(val):
        """文字列に初診・再診が含まれるか判定。(is_shoshin, is_saishin) を返す"""
        s = str(val or "")
        return ("初診" in s), ("再診" in s)

    # ① 初診・再診 列
    type_val = str(row.get(COL_TYPE) or "").strip()
    if "初診" in type_val:
        return "初診"
    if "再診" in type_val:
        return "再診"

    # ② メニュー導線 列
    menu_val = str(row.get(COL_MENU) or "")
    is_shoshin = ("LP" in menu_val
                  or "はじめてのDAILY" in menu_val
                  or "初診" in menu_val)
    is_saishin = "再診" in menu_val
    if is_shoshin and not is_saishin:
        return "初診"
    if is_saishin and not is_shoshin:
        return "再診"

    # ③ メニュー内容 列
    mc_val = str(row.get(COL_MENU_CONTENT) or "")
    if re.search(r"当院で.+受けたことがある方", mc_val):
        return "再診"
    mc_s, mc_r = check(mc_val)
    if mc_s and not mc_r:
        return "初診"
    if mc_r and not mc_s:
        return "再診"

    return "その他"


def fmt_date(d) -> str:
    """date → 'YYYY/MM/DD'"""
    try:
        return pd.Timestamp(d).strftime("%Y/%m/%d")
    except Exception:
        return str(d)


def build_detail(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["予約日"] = pd.to_datetime(df[COL_START], errors="coerce").dt.date
    df["区分"]   = df.apply(classify_row, axis=1)
    return df


def aggregate_base(df: pd.DataFrame, date_from=None, date_to=None) -> pd.DataFrame:
    """日付×区分の集計（行=日付、列=区分）"""
    records = []
    for date, g in df.groupby("予約日", dropna=False):
        records.append({
            "予約日": fmt_date(date),
            "予約数": len(g),
            "初診":   int((g["区分"] == "初診").sum()),
            "再診":   int((g["区分"] == "再診").sum()),
            "その他": int((g["区分"] == "その他").sum()),
        })
    result = pd.DataFrame(records) if records else pd.DataFrame(columns=["予約日", "予約数", "初診", "再診", "その他"])
    if date_from and date_to:
        all_dates = {fmt_date(date_from + datetime.timedelta(days=i))
                     for i in range((date_to - date_from).days + 1)}
        existing = set(result["予約日"]) if not result.empty else set()
        zero_rows = [{"予約日": d, "予約数": 0, "初診": 0, "再診": 0, "その他": 0}
                     for d in all_dates if d not in existing]
        if zero_rows:
            result = pd.concat([result, pd.DataFrame(zero_rows)], ignore_index=True)
    return result.sort_values("予約日").reset_index(drop=True)


def aggregate_base_transposed(df: pd.DataFrame, date_from=None, date_to=None) -> pd.DataFrame:
    """転置版：行=区分、列=日付"""
    base = aggregate_base(df, date_from, date_to)
    if base.empty:
        return pd.DataFrame()
    t = base.set_index("予約日").T.reset_index()
    t.columns.name = None
    t = t.rename(columns={"index": "区分"})
    # 合計列を追加
    date_cols = [c for c in t.columns if c != "区分"]
    t["合計"] = t[date_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1).astype(int)
    return t


def get_names(detail_df: pd.DataFrame, date_str: str, kubun: str) -> list:
    """指定した日付・区分に該当する来院者名リストを返す"""
    mask = detail_df["予約日"].apply(fmt_date) == date_str
    sub = detail_df[mask]
    if kubun != "予約数（全員）":
        sub = sub[sub["区分"] == kubun]
    return sorted(sub[COL_NAME].fillna("（名前なし）").tolist())


def aggregate_conditions_transposed(
    raw: pd.DataFrame,
    conditions: list,
    date_from=None,
    date_to=None,
) -> pd.DataFrame:
    """施術名が行、日付が列の転置テーブルを返す。
    フィルタなしの生データから条件マッチしてカウント。
    """
    if not conditions:
        return pd.DataFrame()

    df = raw.copy()
    df["予約日"] = pd.to_datetime(df[COL_START], errors="coerce").dt.date

    if date_from and date_to:
        dates = [date_from + datetime.timedelta(days=i) for i in range((date_to - date_from).days + 1)]
    else:
        dates = sorted(df["予約日"].dropna().unique())
    date_labels = [fmt_date(d) for d in dates]

    rows = []
    for cond in conditions:
        mask = df.apply(lambda row, c=cond: matches_condition(row, c), axis=1)
        matched = df[mask]

        row_data = {"施術名": cond["name"]}
        total = 0
        for d, label in zip(dates, date_labels):
            cnt = int((matched["予約日"] == d).sum())
            row_data[label] = cnt
            total += cnt
        row_data["合計"] = total
        rows.append(row_data)

    return pd.DataFrame(rows)


def add_total_row(df: pd.DataFrame) -> pd.DataFrame:
    num_cols = [c for c in df.columns if c != "予約日"]
    total = {"予約日": "合計"}
    for c in num_cols:
        total[c] = df[c].sum()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")



def num_col_config(cols):
    return {c: st.column_config.NumberColumn(c, format="%d") for c in cols}


# ---------------------------------------------------------------------------
# CSVアップロード & 院選択
# ---------------------------------------------------------------------------
top_left, top_right = st.columns([3, 1])
with top_left:
    uploaded = st.file_uploader("CSVファイルをアップロード", type=["csv"],
                                 label_visibility="collapsed")
with top_right:
    selected_clinic = st.selectbox("院を選択", CLINICS, key="selected_clinic")
    st.caption("※ 施術別・条件マスタタブに反映")

_csv_ok = False
raw_df = None
final_df = pd.DataFrame()
after_cancel_df = pd.DataFrame()
detail_df = pd.DataFrame()
range_detail_df = pd.DataFrame()
range_after_cancel_df = pd.DataFrame()
range_from = datetime.date.today()
range_to = datetime.date.today()

if uploaded is not None:
    try:
        raw_df = load_csv(uploaded)
        missing_cols = [c for c in REQUIRED_COLS if c not in raw_df.columns]
        if missing_cols:
            st.error(f"以下の列が見つかりません: {', '.join(missing_cols)}")
            st.caption(f"ファイルの列: {', '.join(raw_df.columns.tolist())}")
        else:
            final_df, after_cancel_df, steps = run_pipeline(raw_df)
            detail_df = build_detail(final_df)

            st.markdown("#### 処理の内訳")
            step_cols = st.columns(len(steps))
            for col, (label, count, dropped) in zip(step_cols, steps):
                with col:
                    st.metric(
                        label=label,
                        value=f"{count:,} 件",
                        delta=f"-{dropped:,}" if dropped is not None else None,
                        delta_color="inverse",
                    )
            st.divider()

            if final_df.empty:
                st.warning("集計対象のデータがありません。")
            else:
                _csv_ok = True
                _valid = detail_df["予約日"].dropna()
                _min_d = _valid.min() if not _valid.empty else datetime.date.today()
                _max_d = _valid.max() if not _valid.empty else datetime.date.today()

                _dr1, _dr2 = st.columns(2)
                with _dr1:
                    range_from = st.date_input("集計開始日", value=_min_d,
                                                min_value=_min_d, max_value=_max_d, key="range_from")
                with _dr2:
                    range_to = st.date_input("集計終了日", value=_max_d,
                                              min_value=_min_d, max_value=_max_d, key="range_to")

                if range_from and range_to and range_from <= range_to:
                    range_detail_df = detail_df[
                        (detail_df["予約日"] >= range_from) & (detail_df["予約日"] <= range_to)
                    ].copy()
                    _ac = after_cancel_df.copy()
                    _ac["__d"] = pd.to_datetime(_ac[COL_START], errors="coerce").dt.date
                    range_after_cancel_df = (
                        _ac[(_ac["__d"] >= range_from) & (_ac["__d"] <= range_to)]
                        .drop(columns=["__d"]).copy()
                    )
                else:
                    range_detail_df = detail_df.copy()
                    range_after_cancel_df = after_cancel_df.copy()

                st.divider()
    except ValueError as e:
        st.error(str(e))

# ---------------------------------------------------------------------------
# タブ
# ---------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "　📊  予約件数　",
    "　🌸  施術別件数　",
    "　📋  CSV一覧　",
    "　⚙️  条件マスタ　",
    "　💴  金額集計　",
])

# ==========================================================================
# Tab1: 予約件数
# ==========================================================================
with tab1:
    if not _csv_ok:
        st.info("CSVをアップロードすると集計結果が表示されます。")
    else:
        base_result     = aggregate_base(range_detail_df, range_from, range_to)
        base_transposed = aggregate_base_transposed(range_detail_df, range_from, range_to)

        # ── 転置テーブル表示（行=区分、列=日付）──
        date_cols_t = [c for c in base_transposed.columns if c not in ("区分", "合計")]
        col_cfg_t = {
            "区分": st.column_config.TextColumn("区分", width="small"),
            "合計": st.column_config.NumberColumn("合計", format="%d", width="small"),
            **{c: st.column_config.NumberColumn(c, format="%d", width="small")
               for c in date_cols_t},
        }
        st.dataframe(
            base_transposed,
            use_container_width=True,
            hide_index=True,
            column_config=col_cfg_t,
        )
        st.download_button(
            "📥 予約件数をCSVで保存",
            data=to_csv_bytes(base_transposed),
            file_name="予約件数.csv",
            mime="text/csv",
        )

        st.divider()

        # ── 来院者名の確認 ──
        st.markdown("##### 来院者名を確認する")
        all_dates = sorted(base_result["予約日"].tolist())
        nc1, nc2 = st.columns(2)
        with nc1:
            sel_date = st.selectbox("日付を選択", all_dates, key="name_date")
        with nc2:
            sel_kubun = st.selectbox("区分を選択",
                                      ["予約数（全員）", "初診", "再診", "その他"],
                                      key="name_kubun")

        if sel_date and sel_kubun:
            names = get_names(range_detail_df, sel_date, sel_kubun)
            st.caption(f"{sel_date}　{sel_kubun}　{len(names):,} 人")
            if names:
                name_df = pd.DataFrame({"来院者名": names})
                st.dataframe(name_df, use_container_width=True, hide_index=True,
                             height=min(35 * len(names) + 38, 400))
            else:
                st.info("該当者がいません。")

        # ── 「その他」の内訳確認 ──
        other_df = range_detail_df[range_detail_df["区分"] == "その他"]
        if not other_df.empty:
            st.divider()
            with st.expander(f"「その他」 {len(other_df):,} 件 ― 内訳を確認する"):
                st.caption(
                    "初診・再診どちらにも分類できなかった行です。"
                    "「初診・再診」列・「メニュー導線」列・「メニュー内容」列の値を確認し、"
                    "データまたは条件マスタを修正してください。"
                )
                other_show_cols = [c for c in [
                    COL_NAME, "予約日", COL_TYPE, COL_MENU, COL_MENU_CONTENT
                ] if c in other_df.columns]
                st.dataframe(
                    other_df[other_show_cols].sort_values("予約日").reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                    height=min(35 * len(other_df) + 38, 400),
                )

# ==========================================================================
# Tab2: 集計施術別
# ==========================================================================
with tab2:
    conditions = st.session_state["all_conditions"].get(selected_clinic, [])
    st.markdown(f"#### {selected_clinic}")

    if not conditions:
        st.info(f"{selected_clinic} の条件が登録されていません。「⚙️ 条件マスタ」タブから条件を追加してください。")
    else:
        name_to_cond = {c["name"]: c for c in conditions}

        def _norm(s):
            s = unicodedata.normalize("NFKC", str(s)).strip()
            s = s.lower()
            # 全角・半角スペースをすべて除去
            s = re.sub(r'[\s\u3000]+', '', s)
            # × (U+00D7) と ✕ (U+2715) を x に統一
            s = s.replace('\u00d7', 'x').replace('\u2715', 'x')
            # ハイフン・長音・ダッシュ類を - に統一
            s = re.sub(r'[\u30fc\u2015\u2212\u2013\u2014\u2010\u2011\u2012\u301c\uff5e\u3030~]', '-', s)
            return s

        norm_to_cond = {_norm(k): v for k, v in name_to_cond.items()}

        # ── 施術名入力エリア ──
        st.caption("条件マスタの「表示名」を貼り付けると、その順番で件数を集計します。")

        saved_names = st.session_state["custom_cond_order"].get(selected_clinic, [])
        init_rows = saved_names if saved_names else [""] * 5
        input_df = pd.DataFrame({"施術名": init_rows})

        edited_df = st.data_editor(
            input_df,
            num_rows="dynamic",
            column_config={
                "施術名": st.column_config.TextColumn(
                    "施術名（条件マスタの表示名を貼り付け）", width="large"
                )
            },
            use_container_width=True,
            hide_index=True,
            key=f"cond_editor_{selected_clinic}",
        )

        c1, c2 = st.columns([1, 1])
        run_custom = c1.button("集計", key="run_custom", use_container_width=True)
        clear_all  = c2.button("クリア", key="clear_conds", use_container_width=True)

        if clear_all:
            st.session_state["custom_cond_order"][selected_clinic] = []
            save_custom_order(st.session_state["custom_cond_order"])
            st.session_state["cond_not_found"] = []
            st.rerun()

        if run_custom:
            raw_pasted = [str(n).strip() for n in edited_df["施術名"].tolist() if str(n).strip()]
            matched, not_found = [], []
            for n in raw_pasted:
                cond = norm_to_cond.get(_norm(n))
                if cond:
                    matched.append(cond["name"])
                else:
                    not_found.append(n)
            st.session_state["cond_not_found"] = not_found
            if matched:
                st.session_state["custom_cond_order"][selected_clinic] = matched
                save_custom_order(st.session_state["custom_cond_order"])
                st.rerun()
            else:
                st.info("有効な施術名がありません。条件マスタの「表示名」を貼り付けてください。")

        if st.session_state["cond_not_found"]:
            st.warning(f"条件マスタに見つかりません: {', '.join(st.session_state['cond_not_found'])}")

        # ── 集計結果表示 ──
        ordered_names = st.session_state["custom_cond_order"].get(selected_clinic, [])
        ordered_conds = [name_to_cond[n] for n in ordered_names if n in name_to_cond]

        if not ordered_conds:
            st.info("上の入力欄に表示名を貼り付けて「集計」を押してください。")
        elif not _csv_ok:
            st.info("CSVをアップロードすると集計結果が表示されます。")
        else:
            cond_result = aggregate_conditions_transposed(range_after_cancel_df, ordered_conds, range_from, range_to)

            display2 = cond_result.copy()
            date_cols = [c for c in display2.columns if c not in ("施術名", "合計")]
            for c in date_cols + ["合計"]:
                display2[c] = display2[c].apply(lambda x: "" if x == 0 else x)

            col_config2 = {
                "施術名": st.column_config.TextColumn("施術名", width="large"),
                "合計":   st.column_config.TextColumn("合計", width="small"),
                **{c: st.column_config.TextColumn(c, width="small") for c in date_cols},
            }

            st.dataframe(
                display2,
                use_container_width=True,
                hide_index=True,
                column_config=col_config2,
            )
            st.download_button(
                "📥 集計施術別をCSVで保存",
                data=to_csv_bytes(cond_result),
                file_name="集計施術別.csv",
                mime="text/csv",
            )

# ==========================================================================
# Tab3: 明細一覧
# ==========================================================================
with tab3:
    if not _csv_ok:
        st.info("CSVをアップロードすると明細が表示されます。")
    else:
        st.caption("キャンセル除外・重複正規化後の全件一覧。区分フィルターで「その他」に絞ると手動振り分けに使えます。")

        # フィルター行1：来院者名 + 区分
        f1, f2, f3 = st.columns([3, 1, 1])
        with f1:
            search_text = st.text_input("🔍 来院者名で検索", placeholder="名前の一部を入力")
        with f2:
            kubun_filter = st.selectbox("区分", ["全件", "初診", "再診", "その他"])
        with f3:
            st.write("")  # spacer

        # フィルター行2：日付
        d1, d2 = st.columns(2)
        valid_dates = detail_df["予約日"].dropna()
        min_date = valid_dates.min() if not valid_dates.empty else datetime.date.today()
        max_date = valid_dates.max() if not valid_dates.empty else datetime.date.today()
        with d1:
            date_from = st.date_input("開始日", value=min_date,
                                       min_value=min_date, max_value=max_date)
        with d2:
            date_to = st.date_input("終了日", value=max_date,
                                     min_value=min_date, max_value=max_date)

        filtered = detail_df.copy()
        if search_text.strip():
            filtered = filtered[
                filtered[COL_NAME].fillna("").str.contains(search_text.strip(), na=False)
            ]
        if kubun_filter != "全件":
            filtered = filtered[filtered["区分"] == kubun_filter]
        if date_from and date_to:
            filtered = filtered[
                (filtered["予約日"] >= date_from) & (filtered["予約日"] <= date_to)
            ]

        # 「その他」絞り込み時は判定に使った列を先頭に出して確認しやすく
        priority_cols = ["予約日", COL_START, COL_NAME, "区分", COL_TYPE, COL_MENU]
        other_cols = [
            c for c in raw_df.columns
            if c not in priority_cols + [COL_CANCEL]
            and c in filtered.columns
        ]
        display_cols = [c for c in priority_cols if c in filtered.columns] + other_cols

        st.caption(f"{len(filtered):,} 件表示中")
        st.dataframe(
            filtered[display_cols].sort_values(["予約日", COL_START]).reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
            column_config={
                "区分":   st.column_config.TextColumn("区分", width="small"),
                COL_TYPE: st.column_config.TextColumn("初診・再診（元データ）", width="medium"),
                COL_MENU: st.column_config.TextColumn("メニュー導線", width="large"),
            },
        )
        st.download_button(
            "📥 明細（絞り込み結果）をCSVで保存",
            data=to_csv_bytes(filtered[display_cols].sort_values(["予約日", COL_START])),
            file_name="予約明細.csv",
            mime="text/csv",
        )

# ==========================================================================
# Tab4: 条件マスタ
# ==========================================================================
with tab4:
    clinic_name = selected_clinic
    st.markdown(f"#### {clinic_name}　条件マスタ")
    st.caption("院の切り替えは画面上部の「院を選択」で行ってください。")

    conditions = sorted(
        st.session_state["all_conditions"].get(clinic_name, []),
        key=lambda c: (
            c.get("campaign_end") is None,
            c.get("campaign_end") or ""
        )
    )

    st.markdown("##### 登録済み条件")

    if not conditions:
        st.info("条件が登録されていません。下のフォームから追加してください。")
    else:
        search_q = st.text_input(
            "🔍 検索", placeholder="表示名・カテゴリー・施術名で絞り込む", key="search_cond"
        )

        q = search_q.strip().lower() if search_q else ""
        filtered_conds = [
            c for c in conditions
            if not q or q in c["name"].lower()
               or q in (c.get("category") or "").lower()
               or q in (c.get("treatment") or "").lower()
        ] if q else conditions

        st.caption(f"{len(filtered_conds)} / {len(conditions)} 件表示中")

        h = st.columns([0.4, 2, 2.5, 2.5, 1.5, 0.8])
        for col, label in zip(h, ["", "表示名", "施術カテゴリー条件", "施術名条件", "CP終了日", ""]):
            col.markdown(f"<span style='color:#C04F6A;font-weight:700;font-size:13px;'>{label}</span>",
                         unsafe_allow_html=True)
        st.divider()

        if clinic_name not in st.session_state["selected_ids"]:
            st.session_state["selected_ids"][clinic_name] = set()

        for cond in filtered_conds:
            cid = cond["id"]

            if st.session_state["edit_id"] == cid:
                with st.form(key=f"edit_form_{cid}"):
                    st.markdown(f"**「{cond['name']}」を編集**")
                    new_name  = st.text_input("表示名 *", value=cond["name"])
                    ec1, ec2  = st.columns(2)
                    new_cat   = ec1.text_input("施術カテゴリー条件", value=cond.get("category", ""))
                    new_treat = ec2.text_input("施術名条件", value=cond.get("treatment", ""))
                    current_end = cond.get("campaign_end")
                    try:
                        default_end = datetime.date.fromisoformat(str(current_end)) if current_end else None
                    except Exception:
                        default_end = None
                    new_end = st.date_input("キャンペーン終了日（空白=なし）", value=default_end)
                    sb, cb = st.columns(2)
                    submitted = sb.form_submit_button("保存", use_container_width=True)
                    cancelled = cb.form_submit_button("キャンセル", use_container_width=True)

                if submitted:
                    if not new_name.strip():
                        st.error("表示名は必須です。")
                    else:
                        updated = {
                            "id":           cid,
                            "name":         new_name.strip(),
                            "category":     new_cat.strip(),
                            "treatment":    new_treat.strip(),
                            "campaign_end": new_end.isoformat() if new_end else None,
                        }
                        update_condition(updated)
                        for c in conditions:
                            if c["id"] == cid:
                                c.update(updated)
                        st.session_state["all_conditions"][clinic_name] = conditions
                        st.session_state["edit_id"] = None
                        st.rerun()
                if cancelled:
                    st.session_state["edit_id"] = None
                    st.rerun()
            else:
                row_cols = st.columns([0.4, 2, 2.5, 2.5, 1.5, 0.8])
                is_checked = cid in st.session_state["selected_ids"][clinic_name]
                checked = row_cols[0].checkbox("", value=is_checked, key=f"chk_{clinic_name}_{cid}",
                                               label_visibility="collapsed")
                if checked:
                    st.session_state["selected_ids"][clinic_name].add(cid)
                else:
                    st.session_state["selected_ids"][clinic_name].discard(cid)
                ce = cond.get("campaign_end")
                is_expired = False
                if ce:
                    try:
                        d = datetime.date.fromisoformat(str(ce))
                        ce_str = f"{d.month}/{d.day}まで"
                        is_expired = d < datetime.date.today()
                    except Exception:
                        ce_str = str(ce)
                else:
                    ce_str = "─"
                if is_expired:
                    row_cols[1].markdown(
                        f"<span style='color:#B03060;'>⚠ {cond['name']}</span>",
                        unsafe_allow_html=True)
                    row_cols[2].markdown(
                        f"<span style='color:#B03060;'>{cond.get('category', '') or '（指定なし）'}</span>",
                        unsafe_allow_html=True)
                    row_cols[3].markdown(
                        f"<span style='color:#B03060;'>{cond.get('treatment', '') or '（指定なし）'}</span>",
                        unsafe_allow_html=True)
                    row_cols[4].markdown(
                        f"<span style='background:#FFD6DA;padding:2px 6px;border-radius:4px;"
                        f"font-weight:600;color:#B03060;'>期限切れ {ce_str}</span>",
                        unsafe_allow_html=True)
                else:
                    row_cols[1].write(cond["name"])
                    row_cols[2].write(cond.get("category", "") or "（指定なし）")
                    row_cols[3].write(cond.get("treatment", "") or "（指定なし）")
                    row_cols[4].write(ce_str)
                if row_cols[5].button("編集", key=f"edit_btn_{clinic_name}_{cid}"):
                    st.session_state["edit_id"] = cid
                    st.rerun()

            st.divider()

        # 選択削除ボタン（一覧の下）
        selected = st.session_state["selected_ids"].get(clinic_name, set())
        if st.button(f"🗑 選択削除（{len(selected)}件）", key="sel_del",
                     disabled=(len(selected) == 0)):
            for cid in list(selected):
                delete_condition(cid)
            conditions = [c for c in conditions if c["id"] not in selected]
            st.session_state["all_conditions"][clinic_name] = conditions
            st.session_state["selected_ids"][clinic_name] = set()
            st.rerun()

    # ── 新規追加 ──
    st.markdown("#### 新規追加")
    with st.form(f"add_form_{clinic_name}", clear_on_submit=True):
        add_name  = st.text_input("表示名 *", placeholder="例: ルメッカ")
        b1, b2    = st.columns(2)
        add_cat   = b1.text_input("施術カテゴリー条件", placeholder="空白=すべて対象")
        add_treat = b2.text_input("施術名条件", placeholder="空白=すべて対象")
        add_end   = st.date_input("キャンペーン終了日（空白=なし）", value=None)
        add_submit = st.form_submit_button("追加")

    if add_submit:
        if not add_name.strip():
            st.error("表示名は必須です。")
        else:
            _existing = st.session_state["all_conditions"].get(clinic_name, [])
            _next_order = max((c.get("sort_order", 0) for c in _existing), default=0) + 1
            new_cond = {
                "id":           str(uuid.uuid4()),
                "name":         add_name.strip(),
                "category":     add_cat.strip(),
                "treatment":    add_treat.strip(),
                "sort_order":   _next_order,
                "campaign_end": add_end.isoformat() if add_end else None,
            }
            add_condition(clinic_name, new_cond)
            conditions_list = st.session_state["all_conditions"].get(clinic_name, [])
            conditions_list.append(new_cond)
            st.session_state["all_conditions"][clinic_name] = conditions_list
            st.success(f"「{add_name.strip()}」を追加しました。")
            st.rerun()

    # ── 条件テスト（追加前プレビュー） ──
    st.divider()
    with st.expander("🔍 条件テスト（追加前に確認）"):
        st.caption(
            "カテゴリーと施術名条件を入力すると、現在のCSVで何件マッチするか確認できます。"
            "マスタに登録する前の動作確認にお使いください。"
        )
        prev_cat   = st.text_input("施術カテゴリー条件（テスト）",
                                    placeholder="空白=すべて対象", key="prev_cat")
        prev_treat = st.text_input("施術名条件（テスト）",
                                    placeholder="空白=すべて対象", key="prev_treat")
        if prev_cat.strip() or prev_treat.strip():
            _preview_cond = {
                "category":  prev_cat.strip(),
                "treatment": prev_treat.strip(),
                "name":      "テスト",
            }
            _pmask = after_cancel_df.apply(
                lambda r: matches_condition(r, _preview_cond), axis=1
            )
            _matched = after_cancel_df[_pmask]
            st.info(f"マッチ件数：**{len(_matched):,} 件**")
            if not _matched.empty:
                _pcols = [c for c in [COL_NAME, COL_START, COL_CATEGORY, COL_TREAT]
                          if c in _matched.columns]
                st.dataframe(
                    _matched[_pcols].head(20).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                )

    # ── CSV実際の値を確認（診断） ──
    st.divider()
    st.markdown("#### CSVの実際の値を確認")
    st.caption("マスタに入力する文字列はここからコピーしてください。スペースや記号のズレが件数不一致の原因です。")

    has_cat   = COL_CATEGORY in raw_df.columns
    has_treat = COL_TREAT    in raw_df.columns

    if not has_cat and not has_treat:
        st.warning(f"「{COL_CATEGORY}」「{COL_TREAT}」列がCSVに見つかりません。")
    else:
        d1, d2 = st.columns(2)
        if has_cat:
            with d1:
                st.markdown("**施術カテゴリー** の一覧")
                cats = (
                    raw_df[COL_CATEGORY]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .replace("", float("nan"))
                    .dropna()
                    .unique()
                )
                cats_sorted = sorted(cats)
                st.dataframe(
                    pd.DataFrame({"施術カテゴリー": cats_sorted}),
                    use_container_width=True,
                    hide_index=True,
                    height=300,
                )
        if has_treat:
            with d2:
                st.markdown("**施術名** の一覧")
                filter_cat = ""
                if has_cat:
                    filter_cat = st.selectbox(
                        "カテゴリーで絞り込む",
                        ["（すべて）"] + cats_sorted,
                        key="diag_cat_filter",
                    )
                treats_df = raw_df[[COL_CATEGORY, COL_TREAT]].dropna(subset=[COL_TREAT])
                treats_df = treats_df[treats_df[COL_TREAT].astype(str).str.strip() != ""]
                if filter_cat and filter_cat != "（すべて）":
                    treats_df = treats_df[
                        treats_df[COL_CATEGORY].astype(str).str.strip() == filter_cat
                    ]
                treat_list = sorted(treats_df[COL_TREAT].astype(str).str.strip().unique())
                st.dataframe(
                    pd.DataFrame({"施術名": treat_list}),
                    use_container_width=True,
                    hide_index=True,
                    height=300,
                )

# ==========================================================================
# Tab5: 金額集計
# ==========================================================================
with tab5:
    st.markdown(f"#### 予約金額集計　（{selected_clinic}）")

    _menu_file      = _menu_price_file(selected_clinic)
    _menu_name_file = _menu_price_name_file(selected_clinic)

    # ── マスタ（院ごとにキャッシュ保存） ──
    master_df = None
    if os.path.exists(_menu_file):
        try:
            master_df = pd.read_csv(_menu_file, encoding="utf-8-sig", dtype=str)
        except Exception:
            master_df = None

    if master_df is not None:
        _saved_name = open(_menu_name_file).read().strip() if os.path.exists(_menu_name_file) else os.path.basename(_menu_file)
        mc1, mc2 = st.columns([4, 1])
        mc1.success(f"マスタ読み込み済み：{_saved_name}　{len(master_df):,} 件")
        if mc2.button("✕ 変更する", key="del_menu_csv", use_container_width=True):
            st.session_state["confirm_del_menu_csv"] = True
        if st.session_state.get("confirm_del_menu_csv"):
            st.warning("メニュー情報CSVを削除します。よろしいですか？")
            cc1, cc2 = st.columns(2)
            if cc1.button("はい、削除する", key="confirm_del_yes", type="primary", use_container_width=True):
                os.remove(_menu_file)
                if os.path.exists(_menu_name_file):
                    os.remove(_menu_name_file)
                st.session_state.pop("confirm_del_menu_csv", None)
                st.rerun()
            if cc2.button("キャンセル", key="confirm_del_no", use_container_width=True):
                st.session_state.pop("confirm_del_menu_csv", None)
                st.rerun()
    else:
        csv_file = st.file_uploader(
            "📋 メニュー情報CSVをアップロード",
            type=["csv"], key="menu_price_csv"
        )
        if csv_file is not None:
            try:
                _tmp = pd.read_csv(csv_file, encoding="utf-8-sig", dtype=str)
                if "メニュー名※必須" not in _tmp.columns:
                    st.error("「メニュー名※必須」列が見つかりません。")
                elif "金額" not in _tmp.columns:
                    st.error("「金額」列が見つかりません。")
                else:
                    _tmp = _tmp[_tmp["メニュー名※必須"].notna() & (_tmp["メニュー名※必須"].astype(str).str.strip() != "")]
                    _tmp.to_csv(_menu_file, index=False, encoding="utf-8-sig")
                    with open(_menu_name_file, "w", encoding="utf-8") as f:
                        f.write(csv_file.name)
                    st.rerun()
            except Exception as e:
                st.error(f"読み込みエラー：{e}")
        else:
            st.info("（任意）メニュー情報CSVをアップロードすると、メニュー名と照合して金額を取得します。未設定の場合は備考列から金額を取得します。")

    # 集計条件（Tab1の期間を流用）
    price_from = range_from
    price_to   = range_to
    st.caption(f"集計期間：{range_from}　〜　{range_to}　（予約件数タブの期間と同じ）")
    if master_df is None:
        st.caption("メニューCSV未設定 → 備考列から金額を取得します")
    if not uploaded:
        st.info("先に予約CSVをアップロードしてください（画面上部）。")
    else:
        with st.spinner("集計中..."):
            price_dict  = build_price_dict_from_master(master_df) if master_df is not None else {}
            treat_dict  = build_treatment_dict_from_csv(master_df) if master_df is not None else {}
            detail_df_price, unmatched_df_price, empty_cnt_price = aggregate_prices(
                raw_df, price_dict,
                price_from, price_to,
                treat_csv_dict=treat_dict,
                clinic_name=selected_clinic,
            )

        if detail_df_price.empty:
            st.warning("指定期間に該当する予約がありません。")
        else:
            total_amt   = int(detail_df_price["予約合計金額"].sum())
            total_cnt   = len(detail_df_price)
            unmatch_cnt = len(unmatched_df_price)
            empty_cnt   = empty_cnt_price
            empty_df_matched = detail_df_price[
                (detail_df_price["元メニュー内容"] == "") &
                (detail_df_price["予約合計金額"] > 0)
            ]
            empty_matched_cnt = len(empty_df_matched)
            empty_zero_cnt    = empty_cnt - empty_matched_cnt

            sm1, sm2, sm3 = st.columns(3)
            sm1.metric("予約合計金額", f"¥{total_amt:,}")
            sm2.metric("集計予約件数", f"{total_cnt:,} 件")
            sm3.metric("メニュー空欄", f"{empty_cnt:,} 件")

            pt1, pt2, pt3 = st.tabs(["月別", "日別", "明細"])

            with pt1:
                st.markdown("##### 月別集計")
                df_m = detail_df_price.copy()
                df_m["月"] = pd.to_datetime(df_m["予約開始日時"]).dt.to_period("M").astype(str)
                monthly = df_m.groupby("月")["予約合計金額"].agg(["sum"]).reset_index()
                monthly.columns = ["月", "合計金額"]
                total_row = pd.DataFrame([{"月": "合計", "合計金額": monthly["合計金額"].sum()}])
                monthly_disp = pd.concat([monthly, total_row], ignore_index=True)
                st.dataframe(monthly_disp, use_container_width=True, hide_index=True)
                st.download_button("📥 月別CSVダウンロード",
                    monthly.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                    "金額月別集計.csv", "text/csv", key="dl_monthly")

            with pt2:
                st.markdown("##### 日別集計")
                df_d = detail_df_price.copy()
                daily = df_d.groupby("予約日")["予約合計金額"].agg(["sum"]).reset_index()
                daily.columns = ["日", "合計金額"]
                all_dates = pd.DataFrame({"日": pd.date_range(price_from, price_to, freq="D").date})
                daily = all_dates.merge(daily, on="日", how="left").fillna(0)
                daily["合計金額"] = daily["合計金額"].astype(int)
                daily["日"] = daily["日"].astype(str)
                daily_t = daily.set_index("日").T
                daily_t["合計"] = daily_t.sum(axis=1)
                daily_t.index.name = "項目"
                st.dataframe(daily_t, use_container_width=True)
                st.download_button("📥 日別CSVダウンロード",
                    daily_t.to_csv(encoding="utf-8-sig").encode("utf-8-sig"),
                    "金額日別集計.csv", "text/csv", key="dl_daily")

            with pt3:
                has_src_col = "照合元" in detail_df_price.columns
                df_menu  = detail_df_price[detail_df_price["照合元"] == "メニュー内容"] if has_src_col else detail_df_price
                df_treat = detail_df_price[detail_df_price["照合元"] == "施術名"] if has_src_col else pd.DataFrame()

                # メニュー内容で一致
                menu_cnt   = len(df_menu)
                menu_total = int(df_menu["予約合計金額"].sum())
                st.markdown(f"##### メニュー内容で照合（{menu_cnt} 件 / ¥{menu_total:,}）")
                show_menu_cols = ["予約ID","予約開始日時","来院者名","元メニュー内容","一致メニュー名","採用金額一覧","照合方法","予約合計金額"]
                show_menu_cols_safe = [c for c in show_menu_cols if c in df_menu.columns]
                if not df_menu.empty:
                    st.dataframe(df_menu[show_menu_cols_safe], use_container_width=True, hide_index=True)
                st.download_button("📥 メニュー内容照合CSVダウンロード",
                    df_menu[show_menu_cols_safe].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                    "金額明細_メニュー内容.csv", "text/csv", key="dl_detail_menu")

                # 施術名で一致
                if not df_treat.empty:
                    treat_cnt   = len(df_treat)
                    treat_total = int(df_treat["予約合計金額"].sum())
                    st.markdown(f"---\n##### 施術名で照合（{treat_cnt} 件 / ¥{treat_total:,}）")
                    show_treat_cols = ["予約ID","予約開始日時","来院者名","施術名","一致メニュー名","採用金額一覧","予約合計金額"]
                    show_treat_cols_safe = [c for c in show_treat_cols if c in df_treat.columns]
                    st.dataframe(df_treat[show_treat_cols_safe], use_container_width=True, hide_index=True)
                    st.download_button("📥 施術名照合CSVダウンロード",
                        df_treat[show_treat_cols_safe].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                        "金額明細_施術名.csv", "text/csv", key="dl_detail_treat")

                # 未取得（施術名パスで金額0）
                empty_df_zero = detail_df_price[
                    (detail_df_price["元メニュー内容"] == "") &
                    (detail_df_price["予約合計金額"] == 0)
                ]
                if not empty_df_zero.empty:
                    st.markdown(f"---\n##### 施術名照合→未取得（{empty_zero_cnt} 件 / 金額不明）")
                    show_zero_cols = ["予約ID","予約開始日時","来院者名","施術名"]
                    show_zero_cols_safe = [c for c in show_zero_cols if c in empty_df_zero.columns]
                    st.dataframe(empty_df_zero[show_zero_cols_safe], use_container_width=True, hide_index=True)

                if not unmatched_df_price.empty:
                    st.markdown(f"---\n##### 未一致メニュー（{unmatch_cnt} 件）")
                    st.dataframe(unmatched_df_price, use_container_width=True, hide_index=True)
