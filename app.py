import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os
import re

st.set_page_config(page_title="Discipline Management System", layout="wide")
st.title("Discipline Integrated Dashboard (2018-2026)")
st.markdown("AI Data Integration System Project Protocol.")

EXCEL_FILE = "data.xlsx"
REQUIRED_COLUMNS = ["번호", "년도", "구분", "소속", "직책", "성명", "징계 사유", "징계종류", "징계일"]

def clean_discipline_type(val):
    if pd.isna(val) or not isinstance(val, str): return "기타"
    val = val.strip()
    for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고사직"]:
        if t in val: return t
    return "기타"

def clean_location_name(val):
    if pd.isna(val) or not isinstance(val, str): return "기타 사업장"
    return re.sub(r'\s*\(.*?\)\s*', '', val.replace('\n', ' ')).strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            excel_obj = pd.ExcelFile(EXCEL_FILE)
            combined_rows = []
            for sheet_name in excel_obj.sheet_names:
                raw_df = excel_obj.parse(sheet_name, header=None)
                hdr_idx = 0
                for idx, row in raw_df.iterrows():
                    row_vals = [str(x).strip() for x in row.values if pd.notna(x)]
                    if any(k in row_vals for k in ["번호", "년도", "성명", "소속"]):
                        hdr_idx = idx
                        break
                sheet_df = excel_obj.parse(sheet_name, skiprows=hdr_idx)
                sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                if sheet_df.empty: continue
                for col in REQUIRED_COLUMNS:
                    if col not in sheet_df.columns:
                        m = [c for c in sheet_df.columns if col in c or c in col]
                        sheet_df[col] = sheet_df[m[0]] if m else ""
                sheet_df = sheet_df.dropna(subset=["성명"], how="all")
                if not sheet_df.empty:
                    combined_rows.append(sheet_df[REQUIRED_COLUMNS])
            if not combined_rows: return get_sample_data()
            total_df = pd.concat(combined_rows, ignore_index=True)
            total_df = total_df[total_df["성명"].astype(str).str.strip() != ""]
            total_df["구분"] = total_df["구분"].fillna("미지정").astype(str).str.strip().replace("", "미지정")
            total_df["소속_정제"] = total_df["소속"].apply(clean_location_name)
            total_df["징계종류_정제"] = total_df["징계종류"].apply(clean_discipline_type)
            total_df["년도"] = pd.to_numeric(total_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            total_df["성명"] = total_df["성명"].fillna("미상").astype(str).str.strip()
            total_df = total_df.sort_values(by=["년도", "징계일"]).reset_index(drop=True)
            total_df["번호"] = total_df.index + 1
            return total_df
        except Exception as e:
            st.error(f"Error: {e}")
            return get_sample_data()
    return get_sample_data()

def get_sample_data():
    return pd.DataFrame([{c: "연동대기" for c in REQUIRED_COLUMNS} | {"년도": 2026, "소속_정제": "테스트", "징계종류_정제": "기타"}])

if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

df = st.session_state.discipline_data

st.sidebar.header("Input Form")
with st.sidebar.form(key="input_form", clear_on_submit=True):
    div_opts = list(df["구분"].dropna().unique()) if not df.empty else ["기본 법인"]
    input_division = st.selectbox("Company Select", div_opts + ["기타"] if "기타" not in div_opts else div_opts)
    input_year = st.number_input("Year", min_value=2018, max_value=2030, value=datetime.date.today().year)
    input_date = st.date_input("Date", datetime.date.today())
    input_dept = st.text_input("Department")
    input_position = st.text_input("Position")
    input_name = st.text_input("Name")
    input_type = st.selectbox("Type", ["견책", "감봉", "정직", "강등", "해고", "경고", "훈계", "권고사직"])
    input_reason = st.text_area("Reason")
    submit_button = st.form_submit_button(label="Add Data")

if submit_button and input_dept and input_name and input_reason:
    new_row = {
        "번호": len(df) + 1, "년도": int(input_year), "구분": input_division, "소속": input_dept,
        "소속_정제": clean_location_name(input_dept), "직책": input_position, "성명": input_name,
        "징계 사유": input_reason, "징계종류": input_type, "징계종류_정제": clean_discipline_type(input_type),
        "징계일": input_date.strftime("%Y-%m-%d")
    }
    st.session_state.discipline_data = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    st.sidebar.success("Success")
    st.rerun()

df = st.session_state.discipline_data

st.subheader("Search Filter")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    with col1: f_year = st.multiselect("Year Filter", options=sorted(list(df["년도"].unique())))
    with col2: f_dept = st.multiselect("Department Filter", options=sorted(list(df
