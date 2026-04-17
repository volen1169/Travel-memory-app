import streamlit as st
import pandas as pd
import gspread
import pycountry
import time
import random
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="🌍 🌍 Travel Memory Dashboard", layout="wide")

SHEET_ID = st.secrets["google_sheets"]["sheet_id"]

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_ALIASES = {
    "Places": ["Places", "สถานที่"],
    "Transport": ["Transport", "การเดินทาง"],
    "Hotels": ["Hotels", "ที่พัก"],
    "Food": ["Food", "อาหารและของกิน"],
    "Packages": ["Packages", "แพ็กเกจและซิม"],
    "Others": ["Others", "ค่าใช้จ่ายอื่นๆ"],
}

EXPECTED_HEADERS = {
    "Places": ["ประเทศ", "เมือง", "วัน-เวลา", "ชื่อทริป"],
    "Transport": ["ประเภท", "สาย", "ราคา", "ไฟลท์", "เวลา", "ชื่อทริป"],
    "Hotels": ["ชื่อโรงแรม", "ประเภทห้อง", "ราคา", "ชื่อทริป"],
    "Food": ["ประเภท", "ชื่อ", "ราคา", "ชื่อทริป"],
    "Packages": ["ประเภท", "ชื่อ", "ราคา", "ชื่อทริป"],
    "Others": ["ประเภท", "ชื่อ", "ราคา", "ชื่อทริป"],
}

DISPLAY_NAMES = {
    "Places": "สถานที่",
    "Transport": "การเดินทาง",
    "Hotels": "ที่พัก",
    "Food": "อาหารและของกิน",
    "Packages": "แพ็กเกจและซิม",
    "Others": "ค่าใช้จ่ายอื่นๆ",
}

SECTION_ICONS = {
    "Places": "📍",
    "Transport": "✈️",
    "Hotels": "🏨",
    "Food": "🍜",
    "Packages": "📶",
    "Others": "💸",
}

COST_SHEETS = ["Transport", "Hotels", "Food", "Packages", "Others"]


def inject_custom_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

        :root {
            --bg: #f6f9fc;
            --panel: rgba(255,255,255,0.94);
            --panel-strong: rgba(255,255,255,0.98);
            --text: #0a2540;
            --muted: #5b6b7f;
            --line: rgba(226,232,240,0.96);
            --shadow: 0 18px 38px rgba(15, 23, 42, 0.055);
            --shadow-hover: 0 22px 55px rgba(69, 104, 255, 0.14);
            --radius-xl: 30px;
            --radius-lg: 24px;
            --radius-md: 20px;
            --accent-a: #7dd3fc;
            --accent-b: #60a5fa;
            --accent-c: #6366f1;
            --accent-d: #8b5cf6;
        }

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(96,165,250,0.12), transparent 24%),
                radial-gradient(circle at top right, rgba(56,189,248,0.08), transparent 20%),
                linear-gradient(180deg, #f8fbff 0%, #f4f7fb 52%, #eef4fb 100%);
            color: var(--text);
        }

        .main .block-container {
            max-width: 1320px;
            padding-top: 1.4rem;
            padding-bottom: 3.8rem;
        }

        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(246,250,255,0.98) 100%);
            border-right: 1px solid var(--line);
            box-shadow: 10px 0 32px rgba(15,23,42,0.04);
        }

        [data-testid="stSidebar"] .block-container {
            padding-top: 1.45rem;
            padding-bottom: 1.45rem;
        }

        .hero-card {
            position: relative;
            overflow: hidden;
            background:
                radial-gradient(circle at 84% 18%, rgba(255,255,255,0.25), transparent 16%),
                radial-gradient(circle at 16% 100%, rgba(255,255,255,0.14), transparent 18%),
                linear-gradient(135deg, #0a2540 0%, #1b3b72 40%, #3259ff 100%);
            border: 1px solid rgba(255,255,255,0.15);
            border-radius: 36px;
            padding: 38px 42px 36px 42px;
            color: white;
            box-shadow: 0 28px 74px rgba(37,99,235,0.20);
            margin-bottom: 1.8rem;
        }

        .hero-card::before {
            content: "";
            position: absolute;
            inset: -20% auto auto -8%;
            width: 360px;
            height: 360px;
            background: radial-gradient(circle, rgba(255,255,255,0.14), transparent 62%);
            pointer-events: none;
        }

        .hero-card::after {
            content: "";
            position: absolute;
            inset: 0;
            background: linear-gradient(120deg, transparent 30%, rgba(255,255,255,0.12) 49%, transparent 69%);
            transform: translateX(-125%);
            animation: stripeShine 7s ease-in-out infinite;
            pointer-events: none;
        }

        @keyframes stripeShine {
            0% { transform: translateX(-125%); }
            52% { transform: translateX(125%); }
            100% { transform: translateX(125%); }
        }

        .hero-kicker {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 14px;
            border-radius: 999px;
            background: rgba(255,255,255,0.12);
            border: 1px solid rgba(255,255,255,0.14);
            font-size: 0.80rem;
            font-weight: 800;
            letter-spacing: 0.02em;
            margin-bottom: 1rem;
            backdrop-filter: blur(8px);
        }

        .hero-title {
            font-size: 3.15rem;
            font-weight: 900;
            line-height: 0.98;
            letter-spacing: -0.055em;
            margin-bottom: 0.65rem;
        }

        .hero-subtitle {
            font-size: 1.04rem;
            line-height: 1.78;
            max-width: 920px;
            color: rgba(255,255,255,0.94);
        }

        .sidebar-card,
        .section-card,
        .panel-card,
        .tm-metric,
        .summary-card,
        .detail-item {
            position: relative;
            overflow: hidden;
            background: var(--panel);
            border: 1px solid var(--line);
            box-shadow: var(--shadow);
            transition: transform 180ms ease, box-shadow 180ms ease, border-color 180ms ease, background 180ms ease;
            backdrop-filter: blur(10px);
        }

        .sidebar-card:hover,
        .section-card:hover,
        .panel-card:hover,
        .tm-metric:hover,
        .summary-card:hover,
        .detail-item:hover {
            transform: translateY(-4px);
            border-color: rgba(99,102,241,0.24);
            box-shadow: var(--shadow-hover);
        }

        .sidebar-card::after,
        .section-card::after,
        .panel-card::after,
        .tm-metric::after,
        .summary-card::after,
        .detail-item::after {
            content: "";
            position: absolute;
            inset: -1px;
            background: linear-gradient(135deg, rgba(125,211,252,0.18), rgba(99,102,241,0.08), rgba(139,92,246,0.18));
            opacity: 0;
            transition: opacity 180ms ease;
            pointer-events: none;
        }

        .sidebar-card:hover::after,
        .section-card:hover::after,
        .panel-card:hover::after,
        .tm-metric:hover::after,
        .summary-card:hover::after,
        .detail-item:hover::after {
            opacity: 1;
        }

        .sidebar-card { border-radius: 24px; padding: 18px 18px 16px 18px; margin-bottom: 16px; }
        .sidebar-title { font-size: 1.65rem; font-weight: 900; letter-spacing: -0.04em; color: var(--text); margin-bottom: 0.45rem; }
        .sidebar-subtitle, .sidebar-list, .section-subtitle, .panel-subtitle, .tm-metric-note { color: var(--muted); }
        .sidebar-check { color: var(--text); font-weight: 800; margin-bottom: 0.65rem; }
        .sidebar-list { padding-left: 1.1rem; margin: 0; line-height: 1.8; }

        .tm-metric {
            border-radius: 28px;
            padding: 24px 24px 20px 24px;
            min-height: 148px;
            margin-bottom: 0.4rem;
        }

        .tm-metric::before,
        .section-card::before,
        .panel-card::before,
        .summary-card::before,
        .detail-item::before {
            content: "";
            position: absolute;
            inset: 0 0 auto 0;
            height: 3px;
            background: linear-gradient(90deg, var(--accent-a) 0%, var(--accent-b) 28%, var(--accent-c) 68%, var(--accent-d) 100%);
            opacity: 0.96;
            z-index: 1;
        }

        .tm-metric-label { color: #475569; font-size: 0.92rem; font-weight: 700; margin-bottom: 0.72rem; position: relative; z-index: 2; }
        .tm-metric-value { color: var(--text); font-size: 2.7rem; font-weight: 900; line-height: 1.0; letter-spacing: -0.055em; margin-bottom: 0.38rem; position: relative; z-index: 2; }
        .tm-metric-note { font-size: 0.92rem; line-height: 1.48; position: relative; z-index: 2; }

        .section-card {
            border-radius: var(--radius-xl);
            padding: 26px 28px 22px 28px;
            margin-bottom: 1.6rem;
        }

        .section-title { color: var(--text); font-size: 1.68rem; font-weight: 900; letter-spacing: -0.04em; margin-bottom: 0.24rem; position: relative; z-index: 2; }
        .section-subtitle { font-size: 0.98rem; line-height: 1.62; position: relative; z-index: 2; }

        .panel-card {
            border-radius: var(--radius-xl);
            padding: 24px 24px 20px 24px;
            min-height: 100%;
            margin-bottom: 1.4rem;
        }

        .panel-title { color: var(--text); font-size: 1.32rem; font-weight: 900; letter-spacing: -0.03em; margin-bottom: 0.28rem; position: relative; z-index: 2; }
        .panel-subtitle { font-size: 0.95rem; line-height: 1.58; margin-bottom: 1.15rem; position: relative; z-index: 2; }

        .summary-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 20px;
            margin-top: 0.65rem;
            position: relative;
            z-index: 2;
        }

        .summary-card {
            border-radius: 24px;
            padding: 22px 22px 18px 22px;
            min-height: 158px;
            background: var(--panel-strong);
        }

        .summary-card-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 0.7rem;
        }

        .summary-card-name { color: var(--text); font-size: 1rem; font-weight: 800; position: relative; z-index: 2; }
        .summary-card-count {
            color: #3155ff;
            background: linear-gradient(180deg, rgba(232,240,255,0.96), rgba(241,245,255,0.96));
            border: 1px solid rgba(147,197,253,0.45);
            border-radius: 999px;
            padding: 6px 11px;
            font-size: 0.8rem;
            font-weight: 800;
            white-space: nowrap;
            position: relative;
            z-index: 2;
        }

        .summary-card-value { color: var(--text); font-size: 1.82rem; font-weight: 900; letter-spacing: -0.05em; line-height: 1.02; margin-bottom: 0.28rem; position: relative; z-index: 2; }
        .summary-card-note { color: var(--muted); font-size: 0.9rem; line-height: 1.52; position: relative; z-index: 2; }

        .list-stack {
            display: grid;
            gap: 20px;
            position: relative;
            z-index: 2;
        }

        .detail-item {
            border-radius: 24px;
            padding: 20px 22px 20px 22px;
            background: var(--panel-strong);
        }

        .detail-main { color: var(--text); font-size: 1.1rem; font-weight: 900; margin-bottom: 0.85rem; letter-spacing: -0.02em; position: relative; z-index: 2; }
        .detail-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 14px 18px;
            position: relative;
            z-index: 2;
        }

        .detail-field-label { color: var(--muted); font-size: 0.76rem; font-weight: 800; margin-bottom: 0.18rem; text-transform: uppercase; letter-spacing: 0.05em; }
        .detail-field-value { color: var(--text); font-size: 0.99rem; font-weight: 600; word-break: break-word; }
        .detail-price { color: var(--text); font-size: 1.06rem; font-weight: 900; }

        .empty-state {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 264px;
            border-radius: 24px;
            background: linear-gradient(180deg, rgba(248,250,252,0.96), rgba(241,245,249,0.94));
            border: 1px dashed rgba(148,163,184,0.38);
            color: var(--muted);
            text-align: center;
            line-height: 1.85;
            font-weight: 600;
            padding: 24px;
            position: relative;
            z-index: 2;
        }

        /* Better spacing between native Streamlit columns */
        div[data-testid="stHorizontalBlock"] > div {
            padding-top: 2px;
            padding-bottom: 2px;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 12px;
            padding-bottom: 0.4rem;
        }

        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: nowrap;
            background: rgba(255,255,255,0.86);
            border-radius: 999px;
            border: 1px solid rgba(226,232,240,0.98);
            color: #334155;
            font-weight: 800;
            padding: 0 18px;
            transition: all 170ms ease;
            box-shadow: 0 8px 18px rgba(15,23,42,0.04);
        }

        .stTabs [data-baseweb="tab"]:hover {
            transform: translateY(-1px);
            box-shadow: 0 12px 22px rgba(37,99,235,0.08);
        }

        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
            color: #3155ff !important;
            border-color: rgba(96,165,250,0.30) !important;
            box-shadow: 0 10px 22px rgba(37,99,235,0.10);
        }

        .stSelectbox label, .stDateInput label, .stTimeInput label,
        .stTextInput label, .stNumberInput label, .stRadio label {
            font-weight: 700 !important;
            color: #334155 !important;
        }

        .stButton > button, .stDownloadButton > button {
            border-radius: 999px;
            border: 1px solid rgba(147,197,253,0.45);
            background: linear-gradient(135deg, #ffffff 0%, #eef4ff 100%);
            color: var(--text);
            font-weight: 800;
            box-shadow: 0 10px 24px rgba(37,99,235,0.08);
            transition: all 170ms ease;
            min-height: 44px;
        }

        .stButton > button:hover, .stDownloadButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 14px 26px rgba(37,99,235,0.14);
            border-color: rgba(96,165,250,0.42);
        }

        /* Smooth chart wrapper styling */
        .chart-shell {
            position: relative;
            min-height: 340px;
            border-radius: 26px;
            background:
                radial-gradient(circle at top right, rgba(99,102,241,0.10), transparent 22%),
                linear-gradient(180deg, rgba(255,255,255,0.98), rgba(247,250,252,0.96));
            border: 1px solid rgba(226,232,240,0.96);
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.8);
            padding: 10px 12px 8px 12px;
            overflow: hidden;
        }

        
        .emoji-pill {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            padding: 10px 14px;
            border-radius: 999px;
            background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(239,246,255,0.96));
            border: 1px solid rgba(191,219,254,0.65);
            box-shadow: 0 10px 24px rgba(37,99,235,0.08);
            margin-bottom: 1rem;
        }

        .emoji-pill img {
            width: 18px;
            height: 18px;
            display: block;
        }

        .emoji-pill span {
            color: var(--text);
            font-size: 0.95rem;
            font-weight: 800;
        }

.chart-shell::before {
            content: "";
            position: absolute;
            inset: 0;
            background: linear-gradient(180deg, rgba(99,102,241,0.04), transparent 38%);
            pointer-events: none;
        }

        @media (max-width: 900px) {
            .hero-title { font-size: 2.3rem; }
            .summary-grid, .detail-grid { grid-template-columns: 1fr; }
            .panel-card, .section-card { padding-left: 20px; padding-right: 20px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_resource(show_spinner=False)
def connect_gsheet():
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPE,
    )
    client = gspread.authorize(credentials)
    return call_with_retry(client.open_by_key, SHEET_ID)


def normalize_dataframe(df: pd.DataFrame, sheet_key: str) -> pd.DataFrame:
    expected = EXPECTED_HEADERS[sheet_key]
    if df.empty:
        return pd.DataFrame(columns=expected)
    df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]
    for col in expected:
        if col not in df.columns:
            df[col] = ""
    return df[expected]


def get_worksheet_map(spreadsheet):
    worksheets = call_with_retry(spreadsheet.worksheets)
    return {ws.title: ws for ws in worksheets}


def find_worksheet(sheet_map: dict, sheet_key: str):
    for title in SHEET_ALIASES[sheet_key]:
        if title in sheet_map:
            return sheet_map[title]
    raise ValueError(f"ไม่พบชีตสำหรับ {sheet_key}. กรุณาตั้งชื่อชีตเป็นหนึ่งในนี้: {', '.join(SHEET_ALIASES[sheet_key])}")


@st.cache_data(show_spinner=False)
def load_all_data():
    spreadsheet = connect_gsheet()
    sheet_map = get_worksheet_map(spreadsheet)
    data = {}
    for key in SHEET_ALIASES:
        ws = find_worksheet(sheet_map, key)
        expected_headers = EXPECTED_HEADERS[key]
        all_values = call_with_retry(ws.get_all_values)
        if not all_values:
            data[key] = pd.DataFrame(columns=expected_headers)
            continue
        header_row_index = None
        for i, row in enumerate(all_values):
            trimmed = [str(x).strip() for x in row[:len(expected_headers)]]
            if trimmed == expected_headers:
                header_row_index = i
                break
        if header_row_index is None:
            data[key] = pd.DataFrame(columns=expected_headers)
            continue
        rows = all_values[header_row_index + 1:]
        cleaned_rows = []
        for row in rows:
            row = row[:len(expected_headers)]
            if len(row) < len(expected_headers):
                row += [""] * (len(expected_headers) - len(row))
            if any(str(cell).strip() for cell in row):
                cleaned_rows.append(row)
        df = pd.DataFrame(cleaned_rows, columns=expected_headers)
        data[key] = normalize_dataframe(df, key)
    return data


def append_row(sheet_key: str, row_values: list):
    spreadsheet = connect_gsheet()
    sheet_map = get_worksheet_map(spreadsheet)
    ws = find_worksheet(sheet_map, sheet_key)
    call_with_retry(ws.append_row, row_values, value_input_option="USER_ENTERED")
    load_all_data.clear()


def to_number(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("฿", "", regex=False)
        .str.replace("บาท", "", regex=False)
        .str.strip()
    )
    cleaned = cleaned.replace({"": None, "nan": None, "None": None})
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def get_trip_names(data_dict: dict) -> list[str]:
    trip_names = set()
    for _, df in data_dict.items():
        if "ชื่อทริป" in df.columns and not df.empty:
            for trip in df["ชื่อทริป"].dropna().astype(str):
                trip = trip.strip()
                if trip:
                    trip_names.add(trip)
    return sorted(trip_names)


def compute_trip_total(data_dict: dict, trip_name: str) -> float:
    total = 0.0
    for key in COST_SHEETS:
        df = data_dict[key]
        if not df.empty:
            filtered = df[df["ชื่อทริป"].astype(str).str.strip() == trip_name]
            total += to_number(filtered["ราคา"]).sum()
    return total


def build_cost_summary(data_dict: dict, trip_name: str) -> pd.DataFrame:
    rows = []
    for key in COST_SHEETS:
        df = data_dict[key]
        filtered = df[df["ชื่อทริป"].astype(str).str.strip() == trip_name] if not df.empty else pd.DataFrame(columns=df.columns)
        amount = to_number(filtered["ราคา"]).sum() if not filtered.empty else 0
        count = len(filtered)
        rows.append({"หมวด": DISPLAY_NAMES[key], "จำนวนรายการ": count, "ยอดรวม": amount})
    return pd.DataFrame(rows)


def format_datetime_for_sheet(date_value, time_value=None) -> str:
    if time_value is None:
        return str(date_value)
    combined = datetime.combine(date_value, time_value)
    return combined.strftime("%Y-%m-%d %H:%M")


@st.cache_data(show_spinner=False)
def get_all_countries():
    countries = []
    for country in pycountry.countries:
        name = getattr(country, "name", None)
        if name:
            countries.append(name)
    return sorted(set(countries))


@st.cache_data(show_spinner=False)
def get_subdivisions_by_country(country_name: str):
    if not country_name:
        return []
    matched_country = None
    for country in pycountry.countries:
        if getattr(country, "name", "") == country_name:
            matched_country = country
            break
    if not matched_country:
        try:
            result = pycountry.countries.search_fuzzy(country_name)
            if result:
                matched_country = result[0]
        except Exception:
            return []
    if not matched_country:
        return []
    alpha_2 = getattr(matched_country, "alpha_2", None)
    if not alpha_2:
        return []
    subdivisions = [sub.name for sub in pycountry.subdivisions if getattr(sub, "country_code", "") == alpha_2]
    return sorted(set(subdivisions))


def reset_city_state(prefix: str):
    for key in [f"city_{prefix}", f"custom_city_{prefix}", f"city_text_{prefix}", f"custom_subdivision_{prefix}"]:
        if key in st.session_state:
            del st.session_state[key]


def render_country_city_dropdown(prefix: str = "default"):
    country_options = get_all_countries() + ["Other / อื่นๆ"]
    selected_country = st.selectbox("ประเทศ", options=country_options, key=f"country_{prefix}", on_change=reset_city_state, args=(prefix,))

    if selected_country == "Other / อื่นๆ":
        return (
            st.text_input("กรอกชื่อประเทศ", key=f"custom_country_{prefix}"),
            st.text_input("กรอกเมือง / จังหวัด / รัฐ", key=f"custom_city_{prefix}"),
        )

    subdivision_options = get_subdivisions_by_country(selected_country)
    if subdivision_options:
        selected_city = st.selectbox("เมือง / จังหวัด / รัฐ", options=subdivision_options + ["Other / อื่นๆ"], key=f"city_{prefix}")
        if selected_city == "Other / อื่นๆ":
            selected_city = st.text_input("กรอกเมือง / จังหวัด / รัฐ", key=f"custom_subdivision_{prefix}")
    else:
        selected_city = st.text_input("เมือง / จังหวัด / รัฐ", key=f"city_text_{prefix}")
    return selected_country, selected_city


def metric_card(label: str, value: str, note: str = ""):
    st.markdown(
        f"""
        <div class="tm-metric">
            <div class="tm-metric-label">{label}</div>
            <div class="tm-metric-value">{value}</div>
            <div class="tm-metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="section-card">
            <div class="section-title">{title}</div>
            <div class="section-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def panel_open(title: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="panel-card">
            <div class="panel-title">{title}</div>
            <div class="panel-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True,
    )


def panel_close():
    st.markdown("</div>", unsafe_allow_html=True)


def render_summary_cards(summary_df: pd.DataFrame):
    html_parts = []
    for _, row in summary_df.iterrows():
        amount = float(to_number(pd.Series([row.get("ยอดรวม", 0)])).iloc[0])
        count = int(to_number(pd.Series([row.get("จำนวนรายการ", 0)])).iloc[0])
        category = str(row.get("หมวด", "")).strip()

        card_html = (
            '<div class="summary-card">'
            '<div class="summary-card-top">'
            f'<div class="summary-card-name">{category}</div>'
            f'<div class="summary-card-count">{count} รายการ</div>'
            '</div>'
            f'<div class="summary-card-value">฿ {amount:,.2f}</div>'
            '<div class="summary-card-note">ยอดรวมของหมวดนี้ในทริปที่เลือก</div>'
            '</div>'
        )
        html_parts.append(card_html)

    st.markdown(
        f'<div class="summary-grid">{"".join(html_parts)}</div>',
        unsafe_allow_html=True,
    )


def render_detail_cards(df: pd.DataFrame, currency_cols: list[str] | None = None):
    currency_cols = currency_cols or []
    if df.empty:
        st.info("ทริปนี้ยังไม่มีข้อมูลในหมวดนี้")
        return

    cards_html = []
    for _, row in df.iterrows():
        row_dict = row.to_dict()
        main_value = ""
        for candidate in ["ชื่อ", "ชื่อโรงแรม", "ประเภท", "ประเทศ", "ชื่อทริป"]:
            val = str(row_dict.get(candidate, "")).strip()
            if val:
                main_value = val
                break
        if not main_value:
            main_value = "รายการ"

        field_parts = []
        for col, raw in row_dict.items():
            if col == "ชื่อทริป":
                continue

            if col in currency_cols:
                amount = float(to_number(pd.Series([raw])).iloc[0])
                value_html = f'<div class="detail-field-value detail-price">฿ {amount:,.2f}</div>'
            else:
                value = str(raw).strip() or "-"
                value_html = f'<div class="detail-field-value">{value}</div>'

            field_parts.append(
                '<div class="detail-field">'
                f'<div class="detail-field-label">{col}</div>'
                f'{value_html}'
                '</div>'
            )

        cards_html.append(
            '<div class="detail-item">'
            f'<div class="detail-main">{main_value}</div>'
            f'<div class="detail-grid">{"".join(field_parts)}</div>'
            '</div>'
        )

    st.markdown(
        f'<div class="list-stack">{"".join(cards_html)}</div>',
        unsafe_allow_html=True,
    )
    st.caption(f"จำนวนรายการ: {len(df)}")


def twemoji_url(emoji: str) -> str:
    codepoints = []
    for ch in emoji:
        cp = ord(ch)
        if cp == 0xFE0F:
            continue
        codepoints.append(f"{cp:x}")
    return "https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/svg/" + "-".join(codepoints) + ".svg"


def render_tab_badge(emoji: str, title: str):
    st.markdown(
        f"""
        <div class="emoji-pill">
            <img src="{twemoji_url(emoji)}" alt="{title}">
            <span>{title}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_info():
    st.sidebar.markdown(
        """
        <div class="sidebar-card">
            <div class="sidebar-title">Travel Memory</div>
            <div class="sidebar-subtitle">จัดเก็บทริป ค่าใช้จ่าย ที่พัก และการเดินทางผ่าน Google Sheets ในหน้าตาที่อ่านง่ายขึ้น</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.sidebar.markdown(
        """
        <div class="sidebar-card">
            <div class="sidebar-check">เช็กเบื้องต้นเมื่อเปิดแอปไม่ขึ้น</div>
            <ol class="sidebar-list">
                <li>sheet_id ถูกต้อง</li>
                <li>ชื่อแต่ละชีตถูกต้อง</li>
                <li>แชร์ให้ service account แล้ว</li>
                <li>เปิด Google Sheets API และ Drive API แล้ว</li>
            </ol>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.sidebar.button("🔄 Refresh data", use_container_width=True):
        load_all_data.clear()
        st.rerun()


def render_top_metrics(data_dict: dict):
    trip_names = get_trip_names(data_dict)
    total_cost = sum(compute_trip_total(data_dict, trip) for trip in trip_names)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("จำนวนทริป", str(len(trip_names)), "ทริปทั้งหมดในระบบ")
    with c2:
        metric_card("สถานที่ทั้งหมด", str(len(data_dict["Places"])), "จุดหมายที่ถูกบันทึก")
    with c3:
        metric_card("รายการเดินทาง", str(len(data_dict["Transport"])), "รวมไฟลท์ รถไฟ และการเดินทาง")
    with c4:
        metric_card("ค่าใช้จ่ายรวม", f"฿ {total_cost:,.2f}", "รวมทุกหมวดค่าใช้จ่าย")


def render_dashboard(data_dict: dict):
    from streamlit.components.v1 import html as st_html

    section_header("📊 Dashboard", "สรุปทริปแบบ Mixpanel / Stripe / Linear ที่อ่านง่าย ดูโปร และเห็นภาพรวมเร็ว")
    trip_names = get_trip_names(data_dict)
    if not trip_names:
        st.warning("ยังไม่มีข้อมูลทริปในระบบ")
        return

    selected_trip = st.selectbox("เลือกชื่อทริป", trip_names)
    total_cost = compute_trip_total(data_dict, selected_trip)
    places_df = data_dict["Places"]
    trip_places = places_df[places_df["ชื่อทริป"].astype(str).str.strip() == selected_trip] if not places_df.empty else pd.DataFrame(columns=places_df.columns)
    summary_df = build_cost_summary(data_dict, selected_trip)

    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        metric_card("ชื่อทริป", selected_trip, "ทริปที่กำลังดู")
    with c2:
        metric_card("จำนวนสถานที่", str(len(trip_places)), "สถานที่ในทริปนี้")
    with c3:
        metric_card("ค่าใช้จ่ายของทริป", f"฿ {total_cost:,.2f}", "รวมทุกหมวดของทริปนี้")

    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

    left, right = st.columns([1.08, 0.92], gap="large")
    with left:
        panel_open("💳 สรุปค่าใช้จ่ายรายหมวด", "การ์ดสรุปแบบ startup dashboard ที่มี hover glow และระยะห่างอ่านง่ายขึ้น")
        render_summary_cards(summary_df)
        panel_close()

    with right:
        panel_open("📈 กราฟค่าใช้จ่าย", "กราฟแท่งโทน gradient พร้อม animation แบบ dashboard ยุคใหม่")
        if summary_df["ยอดรวม"].sum() > 0:
            chart_df = summary_df[summary_df["ยอดรวม"] > 0].copy()
            labels = chart_df["หมวด"].astype(str).tolist()
            values = to_number(chart_df["ยอดรวม"]).tolist()

            chart_html = f"""
            <div class="chart-shell">
              <div id="travel_cost_chart" style="width:100%;height:320px;"></div>
            </div>
            <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
            <script>
              const el = document.getElementById('travel_cost_chart');
              const chart = echarts.init(el);
              const option = {{
                animation: true,
                animationDuration: 900,
                animationEasing: 'cubicOut',
                animationDurationUpdate: 700,
                grid: {{ left: 20, right: 12, top: 16, bottom: 28, containLabel: true }},
                tooltip: {{
                  trigger: 'axis',
                  axisPointer: {{ type: 'shadow' }},
                  backgroundColor: 'rgba(10,37,64,0.94)',
                  borderWidth: 0,
                  textStyle: {{ color: '#fff', fontFamily: 'Inter' }},
                  formatter: function(params) {{
                    const p = params[0];
                    return `<div style="padding:2px 4px;"><b>${{p.name}}</b><br/>฿ ${{Number(p.value).toLocaleString(undefined, {{minimumFractionDigits:2, maximumFractionDigits:2}})}}</div>`;
                  }}
                }},
                xAxis: {{
                  type: 'category',
                  data: {labels!r},
                  axisTick: {{ show: false }},
                  axisLine: {{ lineStyle: {{ color: 'rgba(148,163,184,0.35)' }} }},
                  axisLabel: {{ color: '#5b6b7f', fontWeight: 600, fontFamily: 'Inter' }}
                }},
                yAxis: {{
                  type: 'value',
                  splitLine: {{ lineStyle: {{ color: 'rgba(226,232,240,0.85)' }} }},
                  axisLine: {{ show: false }},
                  axisTick: {{ show: false }},
                  axisLabel: {{
                    color: '#5b6b7f',
                    fontWeight: 600,
                    fontFamily: 'Inter',
                    formatter: function(value) {{ return '฿ ' + Number(value).toLocaleString(); }}
                  }}
                }},
                series: [{{
                  data: {values!r},
                  type: 'bar',
                  barWidth: '46%',
                  itemStyle: {{
                    borderRadius: [12, 12, 4, 4],
                    color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                      {{ offset: 0, color: '#7dd3fc' }},
                      {{ offset: 0.45, color: '#60a5fa' }},
                      {{ offset: 1, color: '#6366f1' }}
                    ]),
                    shadowBlur: 16,
                    shadowColor: 'rgba(99,102,241,0.24)',
                    shadowOffsetY: 8
                  }},
                  emphasis: {{
                    itemStyle: {{
                      shadowBlur: 22,
                      shadowColor: 'rgba(99,102,241,0.32)'
                    }}
                  }}
                }}]
              }};
              chart.setOption(option);
              window.addEventListener('resize', function() {{ chart.resize(); }});
            </script>
            """
            st_html(chart_html, height=340)
        else:
            st.markdown(
                '<div class="empty-state">ทริปนี้ยังไม่มีข้อมูลค่าใช้จ่าย<br>ลองเพิ่มค่าเดินทาง ที่พัก หรือค่าอาหารก่อน</div>',
                unsafe_allow_html=True
            )
        panel_close()

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    section_header("🗂️ รายละเอียดแต่ละหมวด", "แต่ละรายการถูกแสดงเป็น card stack แบบ SaaS แทนตาราง พร้อม emoji ที่อ่านง่ายขึ้น")
    detail_tabs = st.tabs([DISPLAY_NAMES[k] for k in SHEET_ALIASES])

    for tab, key in zip(detail_tabs, SHEET_ALIASES):
        with tab:
            render_tab_badge(SECTION_ICONS[key], DISPLAY_NAMES[key])
            df = data_dict[key]
            filtered = df[df["ชื่อทริป"].astype(str).str.strip() == selected_trip] if not df.empty else pd.DataFrame(columns=df.columns)
            currency_cols = ["ราคา"] if "ราคา" in filtered.columns else []
            render_detail_cards(filtered, currency_cols=currency_cols)


def render_places_form(existing_trip_names: list[str]):
    if st.session_state.pop("places_reset_trip_input", False):
        st.session_state.pop("places_new_trip_name", None)
        st.session_state.pop("places_trip_name", None)

    section_header("เพิ่มข้อมูลสถานที่", "บันทึกประเทศ เมือง วันเวลา และผูกกับชื่อทริป")
    col1, col2 = st.columns(2)
    with col1:
        country, city = render_country_city_dropdown(prefix="places")
    with col2:
        date_value = st.date_input("วันที่", key="places_date")
        time_value = st.time_input("เวลา", value=datetime.now().time(), key="places_time")

    trip_card = st.container()
    with trip_card:
        st.markdown('<div class="section-card" style="padding-bottom: 1rem;">', unsafe_allow_html=True)
        trip_mode = st.radio(
            "เลือกวิธีกรอกชื่อทริป",
            ["เลือกจากทริปเดิม", "สร้างชื่อทริปใหม่"],
            horizontal=True,
            key="places_trip_mode",
        )
        if trip_mode == "เลือกจากทริปเดิม" and existing_trip_names:
            trip_name = st.selectbox("ชื่อทริป", existing_trip_names, key="places_trip_name")
        else:
            trip_name = st.text_input("ชื่อทริปใหม่", key="places_new_trip_name")
        st.markdown('</div>', unsafe_allow_html=True)

    with st.form("places_form", clear_on_submit=True):
        submitted = st.form_submit_button("บันทึกสถานที่", use_container_width=True)
        if submitted:
            trip_name = str(trip_name).strip()
            if not country or not city or not trip_name:
                st.error("กรุณาเลือก ประเทศ เมือง/จังหวัด/รัฐ และชื่อทริป ให้ครบ")
            else:
                append_row("Places", [country, city, format_datetime_for_sheet(date_value, time_value), trip_name])
                reset_city_state("places")
                st.session_state["places_reset_trip_input"] = True
                st.success("บันทึกข้อมูลสถานที่เรียบร้อยแล้ว")
                st.rerun()


def render_transport_form(existing_trip_names: list[str]):
    if st.session_state.pop("transport_reset_trip_input", False):
        st.session_state.pop("transport_new_trip_name", None)
        st.session_state.pop("transport_trip_name", None)

    section_header("เพิ่มข้อมูลการเดินทาง", "เก็บประเภทการเดินทาง ผู้ให้บริการ ราคา และเวลา")
    trip_mode = st.radio(
        "เลือกวิธีกรอกชื่อทริป",
        ["เลือกจากทริปเดิม", "สร้างชื่อทริปใหม่"],
        horizontal=True,
        key="transport_trip_mode",
    )
    if trip_mode == "เลือกจากทริปเดิม" and existing_trip_names:
        trip_name = st.selectbox("ชื่อทริป", existing_trip_names, key="transport_trip_name")
    else:
        trip_name = st.text_input("ชื่อทริปใหม่", key="transport_new_trip_name")

    with st.form("transport_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            travel_type = st.selectbox("ประเภท", ["เครื่องบิน", "รถไฟ", "MRT", "Taxi", "Uber", "รถบัส", "เรือ", "อื่นๆ"])
            line = st.text_input("สาย / ผู้ให้บริการ")
        with col2:
            price = st.number_input("ราคา", min_value=0.0, step=100.0)
            flight_no = st.text_input("ไฟลท์")
        with col3:
            date_value = st.date_input("วันที่เดินทาง")
            time_value = st.time_input("เวลาเดินทาง", key="transport_time", value=datetime.now().time())
        submitted = st.form_submit_button("บันทึกการเดินทาง", use_container_width=True)
        if submitted:
            trip_name = str(trip_name).strip()
            if not line or not trip_name:
                st.error("กรุณากรอก สาย/ผู้ให้บริการ และชื่อทริป")
            else:
                append_row("Transport", [travel_type, line, price, flight_no, format_datetime_for_sheet(date_value, time_value), trip_name])
                st.session_state["transport_reset_trip_input"] = True
                st.success("บันทึกข้อมูลการเดินทางเรียบร้อยแล้ว")
                st.rerun()


def render_hotels_form(existing_trip_names: list[str]):
    section_header("🏨 เพิ่มข้อมูลที่พัก", "บันทึกโรงแรม ประเภทห้อง ราคา และเชื่อมกับทริป")
    with st.form("hotels_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            hotel_name = st.text_input("ชื่อโรงแรม")
            room_type = st.text_input("ประเภทห้อง")
        with col2:
            price = st.number_input("ราคา", min_value=0.0, step=100.0)
            trip_name = st.selectbox("ชื่อทริป", existing_trip_names) if existing_trip_names else st.text_input("ชื่อทริป")
        submitted = st.form_submit_button("บันทึกที่พัก", use_container_width=True)
        if submitted:
            if not hotel_name or not trip_name:
                st.error("กรุณากรอกชื่อโรงแรมและชื่อทริป")
            else:
                append_row("Hotels", [hotel_name, room_type, price, trip_name])
                st.success("บันทึกข้อมูลที่พักเรียบร้อยแล้ว")


def render_simple_cost_form(sheet_key: str, title: str, type_options: list[str], existing_trip_names: list[str], key_suffix: str):
    section_header(title, "บันทึกค่าใช้จ่ายพร้อมชื่อหมวดและผูกกับทริป")
    with st.form(f"{sheet_key}_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            item_type = st.selectbox("ประเภท", type_options, key=f"type_{key_suffix}")
            item_name = st.text_input("ชื่อ", key=f"name_{key_suffix}")
        with col2:
            price = st.number_input("ราคา", min_value=0.0, step=50.0, key=f"price_{key_suffix}")
            trip_name = st.selectbox("ชื่อทริป", existing_trip_names, key=f"trip_{key_suffix}") if existing_trip_names else st.text_input("ชื่อทริป", key=f"trip_text_{key_suffix}")
        submitted = st.form_submit_button("บันทึก", use_container_width=True)
        if submitted:
            if not item_name or not trip_name:
                st.error("กรุณากรอกชื่อรายการและชื่อทริป")
            else:
                append_row(sheet_key, [item_type, item_name, price, trip_name])
                st.success("บันทึกข้อมูลเรียบร้อยแล้ว")


def render_all_tables(data_dict: dict):
    section_header("ดูข้อมูลทุกชีต", "จัดรูปแบบตารางให้อ่านง่ายขึ้น และแยกตามหมวดข้อมูล")
    tabs = st.tabs([DISPLAY_NAMES[k] for k in SHEET_ALIASES])
    for tab, key in zip(tabs, SHEET_ALIASES):
        with tab:
            render_tab_badge(SECTION_ICONS[key], DISPLAY_NAMES[key])
            df = data_dict[key]
            if df.empty:
                st.info("ยังไม่มีข้อมูลในหมวดนี้")
            else:
                currency_cols = ["ราคา"] if "ราคา" in df.columns else []
                display_table(df, currency_cols=currency_cols)


def main():
    inject_custom_css()
    st.markdown(
        """
        <div class="hero-card">
            <div class="hero-kicker">Travel planner • Expense tracker</div><div class="hero-title">Travel Memory Dashboard</div>
            <div class="hero-subtitle">เก็บประวัติทริป ค่าใช้จ่าย การเดินทาง ที่พัก และข้อมูลอื่น ๆ ผ่าน Google Sheets ในรูปแบบที่อ่านง่าย สะอาด และดูทันทีว่าแต่ละทริปใช้งบไปเท่าไร</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_sidebar_info()

    try:
        data_dict = load_all_data()
    except Exception as e:
        if is_quota_error(e):
            st.error("Google Sheets ใช้งานเกินโควต้าชั่วคราว กรุณารอสักครู่แล้วกด Refresh data อีกครั้ง")
        else:
            st.error("เชื่อม Google Sheets ไม่สำเร็จ")
        st.exception(e)
        st.stop()

    render_top_metrics(data_dict)
    st.markdown("<div style='height: 12px'></div>", unsafe_allow_html=True)
    page = st.radio("เมนู", ["Dashboard", "เพิ่มข้อมูล", "ดูข้อมูลทั้งหมด"], horizontal=True)
    existing_trip_names = get_trip_names(data_dict)

    if page == "Dashboard":
        render_dashboard(data_dict)
    elif page == "เพิ่มข้อมูล":
        section_header("เพิ่มข้อมูล", "เลือกหมวดที่ต้องการบันทึก แล้วกรอกข้อมูลในฟอร์มด้านล่าง")
        input_tabs = st.tabs(["สถานที่", "การเดินทาง", "ที่พัก", "อาหารและของกิน", "แพ็กเกจและซิม", "ค่าใช้จ่ายอื่นๆ"])
        with input_tabs[0]:
            render_tab_badge("📍", "สถานที่")
            render_places_form(existing_trip_names)
        with input_tabs[1]:
            render_tab_badge("✈️", "การเดินทาง")
            render_transport_form(existing_trip_names)
        with input_tabs[2]:
            render_tab_badge("🏨", "ที่พัก")
            render_hotels_form(existing_trip_names)
        with input_tabs[3]:
            render_tab_badge("🍜", "อาหารและของกิน")
            render_simple_cost_form("Food", "🍜 เพิ่มข้อมูลอาหารและของกิน", ["ร้านอาหาร", "คาเฟ่", "ของหวาน", "street food", "ของฝาก", "อื่นๆ"], existing_trip_names, "food")
        with input_tabs[4]:
            render_tab_badge("📶", "แพ็กเกจและซิม")
            render_simple_cost_form("Packages", "📶 เพิ่มข้อมูลแพ็กเกจและซิม", ["SIM", "แพ็กเกจทัวร์", "บัตรเดินทาง", "ประกัน", "อื่นๆ"], existing_trip_names, "packages")
        with input_tabs[5]:
            render_tab_badge("💸", "ค่าใช้จ่ายอื่นๆ")
            render_simple_cost_form("Others", "💸 เพิ่มข้อมูลค่าใช้จ่ายอื่นๆ", ["ค่าเข้า", "ประกัน", "ของใช้ส่วนตัว", "ค่าธรรมเนียม", "อื่นๆ"], existing_trip_names, "others")
    else:
        render_all_tables(data_dict)


if __name__ == "__main__":
    main()
