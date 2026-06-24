import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os
import re

st.set_page_config(layout="wide")
st.title("Discipline Integrated Dashboard")

EXCEL_FILE = "data.xlsx"
REQUIRED_COLUMNS = [
    "번호", "년도", "구분", "소속", "직책", 
    "성명", "징계 사유", "징계종류", "징계일"
]

def clean_type(val):
    if pd.isna(val) or not isinstance(val, str): return "기타"
    val = val.strip()
    for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고사직"]:
        if t in val: return t
    return "기타"

def clean_loc(val):
    if pd.isna(val) or not isinstance(val, str): return "기타"
    return re.sub(r'\s*\(.*?\)\s*', '', val.replace('\n', ' ')).strip()

def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=REQUIRED_COLUMNS + ["소속_정제", "징계종류_정제"])
    try:
        excel_obj = pd.ExcelFile(EXCEL_FILE)
        combined_rows = []
        for sheet_name in excel_obj.sheet_names:
            raw_df = excel_obj.parse(sheet_name, header=None)
            hdr_idx = 0
            for idx, row in raw_df.iterrows():
                row_vals = [str(x).strip() for x in row.values if pd.notna(x)]
                if any(k in row_vals for k in ["번호", "년도", "성명"]):
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
        
        if not combined_rows:
            return pd.DataFrame(columns=REQUIRED_COLUMNS + ["소속_정제", "징계종류_정제"])
            
        total_df = pd.concat(combined_rows, ignore_index=True)
        total_df = total_df[total_df["성명"].astype(str).str.strip() != ""]
        total_df["구분"] = total_df["구분"].fillna("미지정").astype(str).str.strip()
        total_df["소속_정제"] = total_df["소속"].apply(clean_loc)
        total_df["징계종류_정제"] = total_df["징계종류"].apply(clean_type)
        total_df["년도"] = pd.to_numeric(total_df["년도"], errors='coerce').fillna(2026).astype(int)
        total_df["성명"] = total_df["성명"].fillna("미상").astype(str).str.strip()
        total_df = total_df.sort_values(by=["년도", "징계일"]).reset_index(drop=True)
        total_df["번호"] = total_df.index + 1
        return total_df
    except:
        return pd.DataFrame(columns=REQUIRED_COLUMNS + ["소속_정제", "징계종류_정제"])

if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_data()

df = st.session_state.discipline_data

st.subheader("Search Filter")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    # [🔥 핵심 수정] 줄이 잘리지 않도록 변수를 위에서 생성 후 한 줄을 아주 짧게 배치했습니다.
    opt_year = sorted(list(df["년도"].unique()))
    opt_dept = sorted(list(df["소속_정제"].unique()))
    opt_name = sorted(list(df["성명"].unique()))
    
    with col1: f_year = st.multiselect("Year", options=opt_year)
    with col2: f_dept = st.multiselect("Department", options=opt_dept)
    with col3: f_name = st.multiselect("Name", options=opt_name)
    
    filtered_df = df.copy()
    if f_year: filtered_df = filtered_df[filtered_df["년도"].isin(f_year)]
    if f_dept: filtered_df = filtered_df[filtered_df["소속_정제"].isin(f_dept)]
    if f_name: filtered_df = filtered_df[filtered_df["성명"].isin(f_name)]
else:
    filtered_df = df

st.markdown("---")
st.subheader("Visualization Chart")
if not filtered_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Total Records by Company")
        st.plotly_chart(px.bar(filtered_df, x="구분", color="구분"), use_container_width=True)
        st.markdown("#### Records by Department")
        st.plotly_chart(px.histogram(filtered_df, x="소속_정제", color="징계종류_정제", barmode="stack"), use_container_width=True)
    with c2:
        st.markdown("#### Yearly Trend Analysis")
        yt = filtered_df.groupby("년도").size().reset_index(name="count")
        st.plotly_chart(px.line(yt, x="년도", y="count", markers=True), use_container_width=True)
        st.markdown("#### Discipline Type Ratio")
        st.plotly_chart(px.pie(filtered_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("No data matches the selected criteria.")

st.markdown("---")
st.subheader("Integrated Data List")
if not filtered_df.empty:
    display_df = filtered_df[[c for c in REQUIRED_COLUMNS if c in filtered_df.columns]]
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: display_df.to_excel(w, index=False)
    st.download_button(label="Download Excel", data=output.getvalue(), file_name="report.xlsx")
else:
    st.info("No data available.")
