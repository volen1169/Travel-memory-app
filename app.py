import streamlit as st
import pandas as pd
import gspread
import pycountry
import time
import random
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="Travel Memory Dashboard", layout="wide")

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
    "Places": "??",
    "Transport": "??",
    "Hotels": "??",
    "Food": "??",
    "Packages": "??",
    "Others": "??",
}

COST_SHEETS = ["Transport", "Hotels", "Food", "Packages", "Others"]


def inject_custom_css():
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(180deg, #f6f8fc 0%, #ffffff 45%, #f7f9fd 100%);
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #f8fbff 0%, #eef4ff 100%);
            border-right: 1px solid rgba(100, 116, 139, 0.12);
        }
        .main .block-container {
            padding-top: 2rem;
            padding-bottom: 2.5rem;
            max-width: 1280px;
        }
        .hero-card {
            background: linear-gradient(135deg, #172554 0%, #1d4ed8 55%, #38bdf8 100%);
            border-radius: 24px;
            padding: 28px 30px;
            color: white;
            box-shadow: 0 16px 40px rgba(37, 99, 235, 0.18);
            margin-bottom: 1.25rem;
        }
        .hero-title {
            font-size: 2.6rem;
            font-weight: 800;
            line-height: 1.1;
            margin-bottom: 0.4rem;
        }
        .hero-subtitle {
            opacity: 0.9;
            font-size: 1.02rem;
        }
        .section-card {
            background: rgba(255,255,255,0.92);
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 22px;
            padding: 18px 20px 12px 20px;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
            margin-bottom: 1rem;
        }
        .tm-metric {
            background: white;
            border: 1px solid rgba(148, 163, 184, 0.16);
            border-radius: 20px;
            padding: 18px 18px 14px 18px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
            min-height: 122px;
        }
        .tm-metric-label {
            color: #475569;
            font-size: 0.95rem;
            font-weight: 600;
            margin-bottom: 0.5rem;
        }
        .tm-metric-value {
            color: #0f172a;
            font-size: 2.1rem;
            font-weight: 800;
            line-height: 1.1;
            margin-bottom: 0.2rem;
        }
        .tm-metric-note {
            color: #64748b;
            font-size: 0.88rem;
        }
        .section-title {
            color: #0f172a;
            font-size: 1.4rem;
            font-weight: 800;
            margin-bottom: 0.2rem;
        }
        .section-subtitle {
            color: #64748b;
            font-size: 0.95rem;
            margin-bottom: 0.75rem;
        }
        .panel-card {
            background: rgba(255,255,255,0.92);
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 22px;
            padding: 18px 18px 12px 18px;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
            height: 100%;
        }
        .panel-title {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            color: #0f172a;
            font-size: 1.1rem;
            font-weight: 800;
            margin-bottom: 0.15rem;
        }
        .panel-subtitle {
            color: #64748b;
            font-size: 0.92rem;
            margin-bottom: 0.95rem;
        }
        .empty-state {
            border-radius: 18px;
            background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
            border: 1px dashed rgba(59, 130, 246, 0.24);
            padding: 22px 18px;
            color: #1e3a8a;
            font-weight: 700;
            min-height: 260px;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid rgba(148,163,184,.16);
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 8px 24px rgba(15,23,42,.05);
            background: white;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
        }
        .stTabs [data-baseweb="tab"] {
            height: 42px;
            border-radius: 999px;
            padding: 0 16px;
            background: rgba(255,255,255,.75);
            border: 1px solid rgba(148,163,184,.16);
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
            color: #1d4ed8;
        }
        .stRadio > div {
            gap: 0.75rem;
        }
        .stButton button, .stDownloadButton button {
            border-radius: 999px;
            border: 0;
            font-weight: 700;
            padding: 0.6rem 1.2rem;
            box-shadow: 0 8px 20px rgba(37,99,235,.16);
        }
        .stForm {
            background: rgba(255,255,255,0.8);
            border: 1px solid rgba(148,163,184,.16);
            border-radius: 22px;
            padding: 1rem 1rem 0.2rem 1rem;
            box-shadow: 0 8px 24px rgba(15,23,42,.05);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def is_quota_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return "quota exceeded" in text or "429" in text or "rate limit" in text


def call_with_retry(func, *args, max_retries: int = 5, base_sleep: float = 1.2, **kwargs):
    last_error = None
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            last_error = e
            if not is_quota_error(e) or attempt == max_retries - 1:
                raise
            sleep_s = base_sleep * (2 ** attempt) + random.uniform(0, 0.5)
            time.sleep(sleep_s)
        except Exception as e:
            last_error = e
            raise
    if last_error:
        raise last_error


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


def display_table(df: pd.DataFrame, currency_cols: list[str] | None = None):
    show_df = df.copy()
    currency_cols = currency_cols or []
    for col in currency_cols:
        if col in show_df.columns:
            show_df[col] = to_number(show_df[col]).map(lambda x: f"฿ {x:,.2f}")
    st.dataframe(show_df, use_container_width=True, hide_index=True)
    st.caption(f"จำนวนรายการ: {len(df)}")


def render_sidebar_info():
    st.sidebar.markdown("## ?? Travel Memory")
    st.sidebar.markdown("จัดเก็บทริป ค่าใช้จ่าย และการเดินทางผ่าน Google Sheets แบบดูง่ายขึ้น")
    st.sidebar.markdown("---")
    st.sidebar.markdown("**เช็กเบื้องต้นเมื่อเปิดแอปไม่ขึ้น**")
    st.sidebar.markdown("1. sheet_id ถูกต้อง\n2. ชื่อแต่ละชีตถูกต้อง\n3. share ให้ service account แล้ว\n4. เปิด Google Sheets API / Drive API แล้ว")
    if st.sidebar.button("?? Refresh data", use_container_width=True):
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
    section_header("?? Dashboard", "สรุปทริปแบบอ่านง่าย พร้อมค่าใช้จ่ายและรายละเอียดแต่ละหมวด")
    trip_names = get_trip_names(data_dict)
    if not trip_names:
        st.warning("ยังไม่มีข้อมูลทริปในระบบ")
        return

    selected_trip = st.selectbox("เลือกชื่อทริป", trip_names)
    total_cost = compute_trip_total(data_dict, selected_trip)
    places_df = data_dict["Places"]
    trip_places = places_df[places_df["ชื่อทริป"].astype(str).str.strip() == selected_trip] if not places_df.empty else pd.DataFrame(columns=places_df.columns)
    summary_df = build_cost_summary(data_dict, selected_trip)

    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("ชื่อทริป", selected_trip, "ทริปที่กำลังดู")
    with c2:
        metric_card("จำนวนสถานที่", str(len(trip_places)), "สถานที่ในทริปนี้")
    with c3:
        metric_card("ค่าใช้จ่ายของทริป", f"฿ {total_cost:,.2f}", "รวมทุกหมวดของทริปนี้")

    left, right = st.columns([1.1, 0.9], gap="large")
    with left:
        panel_open("?? สรุปค่าใช้จ่ายรายหมวด", "ดูจำนวนรายการและยอดรวมของแต่ละหมวดในทริปนี้")
        display_table(summary_df, currency_cols=["ยอดรวม"])
        panel_close()
    with right:
        panel_open("?? กราฟค่าใช้จ่าย", "ช่วยเห็นภาพเร็วว่าหมวดไหนใช้เงินมากที่สุด")
        if summary_df["ยอดรวม"].sum() > 0:
            chart_df = summary_df[summary_df["ยอดรวม"] > 0].set_index("หมวด")
            st.bar_chart(chart_df["ยอดรวม"], use_container_width=True)
        else:
            st.markdown('<div class="empty-state">ทริปนี้ยังไม่มีข้อมูลค่าใช้จ่าย<br>ลองเพิ่มค่าเดินทาง ที่พัก หรือค่าอาหารก่อน</div>', unsafe_allow_html=True)
        panel_close()

    section_header("?? รายละเอียดแต่ละหมวด", "แยกดูข้อมูลของทริปนี้ในแต่ละชีต")
    detail_tabs = st.tabs([f"{SECTION_ICONS[k]} {DISPLAY_NAMES[k]}" for k in SHEET_ALIASES])
    for tab, key in zip(detail_tabs, SHEET_ALIASES):
        with tab:
            df = data_dict[key]
            filtered = df[df["ชื่อทริป"].astype(str).str.strip() == selected_trip] if not df.empty else pd.DataFrame(columns=df.columns)
            if filtered.empty:
                st.info("ทริปนี้ยังไม่มีข้อมูลในหมวดนี้")
            else:
                currency_cols = ["ราคา"] if "ราคา" in filtered.columns else []
                display_table(filtered, currency_cols=currency_cols)


def render_places_form(existing_trip_names: list[str]):
    section_header("?? เพิ่มข้อมูลสถานที่", "บันทึกประเทศ เมือง วันเวลา และผูกกับชื่อทริป")
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
                st.success("บันทึกข้อมูลสถานที่เรียบร้อยแล้ว")
                reset_city_state("places")
                st.session_state["places_new_trip_name"] = ""


def render_transport_form(existing_trip_names: list[str]):
    section_header("?? เพิ่มข้อมูลการเดินทาง", "เก็บประเภทการเดินทาง ผู้ให้บริการ ราคา และเวลา")
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
                st.success("บันทึกข้อมูลการเดินทางเรียบร้อยแล้ว")
                st.session_state["transport_new_trip_name"] = ""


def render_hotels_form(existing_trip_names: list[str]):
    section_header("?? เพิ่มข้อมูลที่พัก", "บันทึกโรงแรม ประเภทห้อง ราคา และเชื่อมกับทริป")
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
    section_header("?? ดูข้อมูลทุกชีต", "จัดรูปแบบตารางให้อ่านง่ายขึ้น และแยกตามหมวดข้อมูล")
    tabs = st.tabs([f"{SECTION_ICONS[k]} {DISPLAY_NAMES[k]}" for k in SHEET_ALIASES])
    for tab, key in zip(tabs, SHEET_ALIASES):
        with tab:
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
            <div class="hero-title">?? Travel Memory Dashboard</div>
            <div class="hero-subtitle">เก็บประวัติทริป ค่าใช้จ่าย การเดินทาง ที่พัก และข้อมูลอื่น ๆ ผ่าน Google Sheets ในรูปแบบที่อ่านง่ายขึ้น</div>
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
        section_header("? เพิ่มข้อมูล", "เลือกหมวดที่ต้องการบันทึก แล้วกรอกข้อมูลในฟอร์มด้านล่าง")
        input_tabs = st.tabs(["?? สถานที่", "?? การเดินทาง", "?? ที่พัก", "?? อาหารและของกิน", "?? แพ็กเกจและซิม", "?? ค่าใช้จ่ายอื่นๆ"])
        with input_tabs[0]:
            render_places_form(existing_trip_names)
        with input_tabs[1]:
            render_transport_form(existing_trip_names)
        with input_tabs[2]:
            render_hotels_form(existing_trip_names)
        with input_tabs[3]:
            render_simple_cost_form("Food", "?? เพิ่มข้อมูลอาหารและของกิน", ["ร้านอาหาร", "คาเฟ่", "ของหวาน", "street food", "ของฝาก", "อื่นๆ"], existing_trip_names, "food")
        with input_tabs[4]:
            render_simple_cost_form("Packages", "?? เพิ่มข้อมูลแพ็กเกจและซิม", ["SIM", "แพ็กเกจทัวร์", "บัตรเดินทาง", "ประกัน", "อื่นๆ"], existing_trip_names, "packages")
        with input_tabs[5]:
            render_simple_cost_form("Others", "?? เพิ่มข้อมูลค่าใช้จ่ายอื่นๆ", ["ค่าเข้า", "ประกัน", "ของใช้ส่วนตัว", "ค่าธรรมเนียม", "อื่นๆ"], existing_trip_names, "others")
    else:
        render_all_tables(data_dict)


if __name__ == "__main__":
    main()
