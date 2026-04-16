import streamlit as st
import pandas as pd
import uuid
import datetime
import unicodedata
import re
import json
from io import BytesIO
from supabase import create_client

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

REQUIRED_COLS = [COL_NAME, COL_START, COL_CANCEL, COL_TYPE]
CLINICS = ["心斎橋院", "新宿院", "福岡院"]


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

if uploaded is None:
    st.markdown(
        "<div style='text-align:center;color:#C04F6A;padding:32px 0;font-size:15px;'>"
        "CSVファイルをアップロードしてください</div>",
        unsafe_allow_html=True,
    )
    st.stop()

try:
    raw_df = load_csv(uploaded)
except ValueError as e:
    st.error(str(e))
    st.stop()

missing_cols = [c for c in REQUIRED_COLS if c not in raw_df.columns]
if missing_cols:
    st.error(f"以下の列が見つかりません: {', '.join(missing_cols)}")
    st.caption(f"ファイルの列: {', '.join(raw_df.columns.tolist())}")
    st.stop()

final_df, after_cancel_df, steps = run_pipeline(raw_df)
detail_df = build_detail(final_df)

# 中間件数
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
    st.stop()

# ---------------------------------------------------------------------------
# 日付範囲フィルター（Tab1・Tab2 共通）
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# タブ
# ---------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "　📊  予約件数　",
    "　🌸  施術別件数　",
    "　📋  CSV一覧　",
    "　⚙️  条件マスタ　",
])

# ==========================================================================
# Tab1: 予約件数
# ==========================================================================
with tab1:
    base_result   = aggregate_base(range_detail_df, range_from, range_to)
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
