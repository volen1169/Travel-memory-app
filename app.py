import streamlit as st
import pandas as pd
import gspread
import pycountry
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

COST_SHEETS = ["Transport", "Hotels", "Food", "Packages", "Others"]


@st.cache_resource(show_spinner=False)
def connect_gsheet():
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPE,
    )
    client = gspread.authorize(credentials)
    return client.open_by_key(SHEET_ID)


def normalize_dataframe(df: pd.DataFrame, sheet_key: str) -> pd.DataFrame:
    expected = EXPECTED_HEADERS[sheet_key]

    if df.empty:
        return pd.DataFrame(columns=expected)

    # ลบคอลัมน์ว่าง
    df = df.loc[:, [c for c in df.columns if str(c).strip() != ""]]

    # เติมคอลัมน์ที่ขาด
    for col in expected:
        if col not in df.columns:
            df[col] = ""

    # เรียงตาม expected
    df = df[expected]

    return df


def find_worksheet(spreadsheet, sheet_key: str):
    aliases = SHEET_ALIASES[sheet_key]
    available_titles = [ws.title for ws in spreadsheet.worksheets()]

    for title in aliases:
        if title in available_titles:
            return spreadsheet.worksheet(title)

    raise ValueError(
        f"ไม่พบชีตสำหรับ {sheet_key}. กรุณาตั้งชื่อชีตเป็นหนึ่งในนี้: {', '.join(aliases)}"
    )


@st.cache_data(ttl=10, show_spinner=False)
def load_all_data():
    spreadsheet = connect_gsheet()
    data = {}

    for key in SHEET_ALIASES:
        ws = find_worksheet(spreadsheet, key)
        expected_headers = EXPECTED_HEADERS[key]
        all_values = ws.get_all_values()

        if not all_values:
            df = pd.DataFrame(columns=expected_headers)
            data[key] = df
            continue

        header_row_index = None

        for i, row in enumerate(all_values):
            trimmed = [str(x).strip() for x in row[:len(expected_headers)]]
            if trimmed == expected_headers:
                header_row_index = i
                break

        # ถ้าไม่เจอ header ที่ตรง ให้ลองใช้แถวแรกแทน
        if header_row_index is None:
            first_row = [str(x).strip() for x in all_values[0][:len(expected_headers)]]
            if first_row == expected_headers:
                header_row_index = 0

        if header_row_index is None:
            df = pd.DataFrame(columns=expected_headers)
            data[key] = df
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
    ws = find_worksheet(spreadsheet, sheet_key)
    ws.append_row(row_values, value_input_option="USER_ENTERED")
    load_all_data.clear()


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


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

        if df.empty:
            amount = 0
            count = 0
        else:
            filtered = df[df["ชื่อทริป"].astype(str).str.strip() == trip_name]
            amount = to_number(filtered["ราคา"]).sum()
            count = len(filtered)

        rows.append({
            "หมวด": DISPLAY_NAMES[key],
            "จำนวนรายการ": count,
            "ยอดรวม": amount,
        })

    return pd.DataFrame(rows)


def format_datetime_for_sheet(date_value, time_value=None) -> str:
    if time_value is None:
        return str(date_value)

    combined = datetime.combine(date_value, time_value)
    return combined.strftime("%Y-%m-%d %H:%M")


def get_all_countries():
    countries = []
    for country in pycountry.countries:
        name = getattr(country, "name", None)
        if name:
            countries.append(name)

    return sorted(set(countries))


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

    subdivisions = [
        sub.name
        for sub in pycountry.subdivisions
        if getattr(sub, "country_code", "") == alpha_2
    ]

    return sorted(set(subdivisions))


def render_country_city_dropdown(prefix: str = "default"):
    country_options = get_all_countries() + ["Other / อื่นๆ"]

    selected_country = st.selectbox(
        "ประเทศ",
        options=country_options,
        key=f"country_{prefix}"
    )

    if selected_country == "Other / อื่นๆ":
        custom_country = st.text_input("กรอกชื่อประเทศ", key=f"custom_country_{prefix}")
        custom_city = st.text_input("กรอกเมือง / จังหวัด / รัฐ", key=f"custom_city_{prefix}")
        return custom_country, custom_city

    subdivision_options = get_subdivisions_by_country(selected_country)

    if subdivision_options:
        selected_city = st.selectbox(
            "เมือง / จังหวัด / รัฐ",
            options=subdivision_options + ["Other / อื่นๆ"],
            key=f"city_{prefix}"
        )

        if selected_city == "Other / อื่นๆ":
            selected_city = st.text_input(
                "กรอกเมือง / จังหวัด / รัฐ",
                key=f"custom_subdivision_{prefix}"
            )
    else:
        selected_city = st.text_input(
            "เมือง / จังหวัด / รัฐ",
            key=f"city_text_{prefix}"
        )

    return selected_country, selected_city


def render_sidebar_info():
    st.sidebar.title("🌍 Travel Memory")
    st.sidebar.info(
        "ระบบนี้เชื่อมกับ Google Sheets แบบหลายชีต\n\n"
        "ถ้าแอปเปิดไม่ขึ้น ให้ตรวจ:\n"
        "1) sheet_id ถูกต้อง\n"
        "2) ชื่อแต่ละชีตถูกต้อง\n"
        "3) share ให้ service account แล้ว\n"
        "4) เปิด Google Sheets API / Drive API แล้ว"
    )

    if st.sidebar.button("🔄 Refresh data"):
        load_all_data.clear()
        st.rerun()


def render_top_metrics(data_dict: dict):
    trip_names = get_trip_names(data_dict)
    total_cost = sum(compute_trip_total(data_dict, trip) for trip in trip_names)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("จำนวนทริป", len(trip_names))
    c2.metric("สถานที่ทั้งหมด", len(data_dict["Places"]))
    c3.metric("รายการเดินทาง", len(data_dict["Transport"]))
    c4.metric("ค่าใช้จ่ายรวมทั้งหมด", f"{total_cost:,.2f} บาท")


def render_dashboard(data_dict: dict):
    st.subheader("📊 Dashboard")

    trip_names = get_trip_names(data_dict)
    if not trip_names:
        st.warning("ยังไม่มีข้อมูลทริปในระบบ")
        return

    selected_trip = st.selectbox("เลือกชื่อทริป", trip_names)

    total_cost = compute_trip_total(data_dict, selected_trip)
    places_df = data_dict["Places"]

    if places_df.empty:
        trip_places = pd.DataFrame(columns=places_df.columns)
    else:
        trip_places = places_df[places_df["ชื่อทริป"].astype(str).str.strip() == selected_trip]

    col1, col2, col3 = st.columns(3)
    col1.metric("ชื่อทริป", selected_trip)
    col2.metric("จำนวนสถานที่", len(trip_places))
    col3.metric("ค่าใช้จ่ายรวมของทริป", f"{total_cost:,.2f} บาท")

    st.markdown("### สรุปค่าใช้จ่ายรายหมวด")
    summary_df = build_cost_summary(data_dict, selected_trip)
    st.dataframe(summary_df, use_container_width=True)

    if summary_df["ยอดรวม"].sum() > 0:
        chart_df = summary_df[summary_df["ยอดรวม"] > 0].set_index("หมวด")
        st.bar_chart(chart_df["ยอดรวม"])

    st.markdown("### รายละเอียดตามชีต")
    detail_tabs = st.tabs([DISPLAY_NAMES[k] for k in SHEET_ALIASES])

    for tab, key in zip(detail_tabs, SHEET_ALIASES):
        with tab:
            df = data_dict[key]

            if df.empty:
                st.info("ยังไม่มีข้อมูล")
            else:
                filtered = df[df["ชื่อทริป"].astype(str).str.strip() == selected_trip]
                if filtered.empty:
                    st.info("ทริปนี้ยังไม่มีข้อมูลในชีตนี้")
                else:
                    st.dataframe(filtered, use_container_width=True)


def render_places_form(existing_trip_names: list[str]):
    st.markdown("### 📍 เพิ่มข้อมูลสถานที่")

    with st.form("places_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            country, city = render_country_city_dropdown(prefix="places")

        with col2:
            date_value = st.date_input("วันที่")
            time_value = st.time_input("เวลา", value=datetime.now().time())

        trip_mode = st.radio(
            "เลือกวิธีกรอกชื่อทริป",
            ["เลือกจากทริปเดิม", "สร้างชื่อทริปใหม่"],
            horizontal=True
        )

        if trip_mode == "เลือกจากทริปเดิม" and existing_trip_names:
            trip_name = st.selectbox("ชื่อทริป", existing_trip_names)
        else:
            trip_name = st.text_input("ชื่อทริปใหม่")

        submitted = st.form_submit_button("บันทึกสถานที่", use_container_width=True)

        if submitted:
            if not country or not city or not trip_name:
                st.error("กรุณาเลือก ประเทศ เมือง/จังหวัด/รัฐ และชื่อทริป ให้ครบ")
            else:
                append_row(
                    "Places",
                    [country, city, format_datetime_for_sheet(date_value, time_value), trip_name],
                )
                st.success("บันทึกข้อมูลสถานที่เรียบร้อยแล้ว")


def render_transport_form(existing_trip_names: list[str]):
    st.markdown("### ✈️ เพิ่มข้อมูลการเดินทาง")

    with st.form("transport_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            travel_type = st.selectbox(
                "ประเภท",
                ["เครื่องบิน", "รถไฟ", "MRT", "Taxi", "Uber", "รถบัส", "เรือ", "อื่นๆ"]
            )
            line = st.text_input("สาย / ผู้ให้บริการ")

        with col2:
            price = st.number_input("ราคา", min_value=0.0, step=100.0)
            flight_no = st.text_input("ไฟลท์")

        with col3:
            date_value = st.date_input("วันที่เดินทาง")
            time_value = st.time_input("เวลาเดินทาง", key="transport_time", value=datetime.now().time())

        trip_mode = st.radio(
            "เลือกวิธีกรอกชื่อทริป ",
            ["เลือกจากทริปเดิม", "สร้างชื่อทริปใหม่"],
            horizontal=True
        )

        if trip_mode == "เลือกจากทริปเดิม" and existing_trip_names:
            trip_name = st.selectbox("ชื่อทริป ", existing_trip_names)
        else:
            trip_name = st.text_input("ชื่อทริปใหม่ ")

        submitted = st.form_submit_button("บันทึกการเดินทาง", use_container_width=True)

        if submitted:
            if not line or not trip_name:
                st.error("กรุณากรอก สาย/ผู้ให้บริการ และชื่อทริป")
            else:
                append_row(
                    "Transport",
                    [travel_type, line, price, flight_no, format_datetime_for_sheet(date_value, time_value), trip_name],
                )
                st.success("บันทึกข้อมูลการเดินทางเรียบร้อยแล้ว")


def render_hotels_form(existing_trip_names: list[str]):
    st.markdown("### 🏨 เพิ่มข้อมูลที่พัก")

    with st.form("hotels_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            hotel_name = st.text_input("ชื่อโรงแรม")
            room_type = st.text_input("ประเภทห้อง")

        with col2:
            price = st.number_input("ราคา ", min_value=0.0, step=100.0)
            if existing_trip_names:
                trip_name = st.selectbox("ชื่อทริป  ", existing_trip_names)
            else:
                trip_name = st.text_input("ชื่อทริป  ")

        submitted = st.form_submit_button("บันทึกที่พัก", use_container_width=True)

        if submitted:
            if not hotel_name or not trip_name:
                st.error("กรุณากรอกชื่อโรงแรมและชื่อทริป")
            else:
                append_row("Hotels", [hotel_name, room_type, price, trip_name])
                st.success("บันทึกข้อมูลที่พักเรียบร้อยแล้ว")


def render_simple_cost_form(
    sheet_key: str,
    title: str,
    type_options: list[str],
    existing_trip_names: list[str],
    key_suffix: str,
):
    st.markdown(f"### {title}")

    with st.form(f"{sheet_key}_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            item_type = st.selectbox("ประเภท", type_options, key=f"type_{key_suffix}")
            item_name = st.text_input("ชื่อ", key=f"name_{key_suffix}")

        with col2:
            price = st.number_input("ราคา", min_value=0.0, step=50.0, key=f"price_{key_suffix}")
            if existing_trip_names:
                trip_name = st.selectbox("ชื่อทริป", existing_trip_names, key=f"trip_{key_suffix}")
            else:
                trip_name = st.text_input("ชื่อทริป", key=f"trip_text_{key_suffix}")

        submitted = st.form_submit_button("บันทึก", use_container_width=True)

        if submitted:
            if not item_name or not trip_name:
                st.error("กรุณากรอกชื่อรายการและชื่อทริป")
            else:
                append_row(sheet_key, [item_type, item_name, price, trip_name])
                st.success("บันทึกข้อมูลเรียบร้อยแล้ว")


def render_all_tables(data_dict: dict):
    st.subheader("📋 ดูข้อมูลทุกชีต")
    tabs = st.tabs([DISPLAY_NAMES[k] for k in SHEET_ALIASES])

    for tab, key in zip(tabs, SHEET_ALIASES):
        with tab:
            df = data_dict[key]
            st.dataframe(df, use_container_width=True)
            st.caption(f"จำนวนรายการ: {len(df)}")


def main():
    st.title("🌍 Travel Memory Dashboard")
    st.caption("เก็บประวัติทริป, ค่าใช้จ่าย, การเดินทาง, ที่พัก และข้อมูลอื่นๆ ผ่าน Google Sheets")

    render_sidebar_info()

    try:
        data_dict = load_all_data()
    except Exception as e:
        st.error("เชื่อม Google Sheets ไม่สำเร็จ")
        st.exception(e)
        st.stop()

    render_top_metrics(data_dict)

    page = st.radio(
        "เมนู",
        ["Dashboard", "เพิ่มข้อมูล", "ดูข้อมูลทั้งหมด"],
        horizontal=True,
    )

    existing_trip_names = get_trip_names(data_dict)

    if page == "Dashboard":
        render_dashboard(data_dict)

    elif page == "เพิ่มข้อมูล":
        input_tabs = st.tabs([
            "สถานที่",
            "การเดินทาง",
            "ที่พัก",
            "อาหารและของกิน",
            "แพ็กเกจและซิม",
            "ค่าใช้จ่ายอื่นๆ",
        ])

        with input_tabs[0]:
            render_places_form(existing_trip_names)

        with input_tabs[1]:
            render_transport_form(existing_trip_names)

        with input_tabs[2]:
            render_hotels_form(existing_trip_names)

        with input_tabs[3]:
            render_simple_cost_form(
                sheet_key="Food",
                title="🍜 เพิ่มข้อมูลอาหารและของกิน",
                type_options=["ร้านอาหาร", "คาเฟ่", "ของหวาน", "street food", "ของฝาก", "อื่นๆ"],
                existing_trip_names=existing_trip_names,
                key_suffix="food",
            )

        with input_tabs[4]:
            render_simple_cost_form(
                sheet_key="Packages",
                title="📶 เพิ่มข้อมูลแพ็กเกจและซิม",
                type_options=["SIM", "แพ็กเกจทัวร์", "บัตรเดินทาง", "ประกัน", "อื่นๆ"],
                existing_trip_names=existing_trip_names,
                key_suffix="packages",
            )

        with input_tabs[5]:
            render_simple_cost_form(
                sheet_key="Others",
                title="💸 เพิ่มข้อมูลค่าใช้จ่ายอื่นๆ",
                type_options=["ค่าเข้า", "ประกัน", "ของใช้ส่วนตัว", "ค่าธรรมเนียม", "อื่นๆ"],
                existing_trip_names=existing_trip_names,
                key_suffix="others",
            )

    elif page == "ดูข้อมูลทั้งหมด":
        render_all_tables(data_dict)


if __name__ == "__main__":
    main()
