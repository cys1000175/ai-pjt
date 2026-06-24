import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re

st.set_page_config(layout="wide")
st.title("HR Dashboard (2018-2026)")

EXCEL_FILE = "data.xlsx"
COLS = ["번호", "년도", "구분", "소속", "직책", "성명", "징계 사유", "징계종류", "징계일"]

def clean_type(val):
    if pd.isna(val) or not isinstance(val, str): return "기타"
    val = val.strip()
    for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고사직", "전배"]:
        if t in val: return t
    return "기타"

def clean_loc(val):
    if pd.isna(val) or not isinstance(val, str): return "미지정"
    return re.sub(r'\s*\(.*?\)\s*', '', val.replace('\n', ' ')).strip()

@st.cache_data
def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])
    try:
        excel_obj = pd.ExcelFile(EXCEL_FILE)
        rows = []
        for s in excel_obj.sheet_names:
            df_s = excel_obj.parse(s, header=None)
            if df_s.empty: continue
            hdr = 0
            for i, r in df_s.iterrows():
                v = [str(x).strip() for x in r.values if pd.notna(x)]
                if any("번호" in x or "소속" in x or "성명" in x for x in v):
                    hdr = i
                    break
            df_real = excel_obj.parse(s, skiprows=hdr)
            df_real.columns = [str(c).strip() for c in df_real.columns]
            df_real = df_real.iloc[:, :11] 
            for c in COLS:
                if c not in df_real.columns:
                    m = [x for x in df_real.columns if c in x or x in c]
                    df_real[c] = df_real[m[0]] if m else ""
            if "징계일" not in df_real.columns or df_real["징계일"].dropna().empty:
                if "일 자" in df_real.columns: df_real["징계일"] = df_real["일 자"]
            rows.append(df_real[COLS])
        if not rows: return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])
        df_tot = pd.concat(rows, ignore_index=True).dropna(subset=["성명", "소속", "징계종류"], how="all")
        df_tot["성명"] = df_tot["성명"].fillna("미상").astype(str).str.strip()
        df_tot = df_tot[df_tot["성명"] != ""]
        df_tot["구분"] = df_tot["구분"].fillna("인사위").astype(str).str.strip().replace("", "인사위")
        df_tot["소속"] = df_tot["소속"].fillna("미지정").astype(str).str.strip()
        df_tot["징계종류"] = df_tot["징계종류"].fillna("기타").astype(str).str.strip()
        df_tot["소속_정제"] = df_tot["소속"].apply(clean_loc)
        df_tot["징계종류_정제"] = df_tot["징계종류"].apply(clean_type)
        df_tot["징계일"] = df_tot["징계일"].astype(str).str.split(" ").str[0]
        df_tot["년도_추출"] = df_tot["징계일"].str.split("-").str[0]
        df_tot["년도"] = pd.to_numeric(df_tot["년도_추출"], errors='coerce').fillna(df_tot["년도"]).fillna(2026)
        df_tot["년도"] = pd.to_numeric(df_tot["년도"], errors='coerce').fillna(2026).astype(int)
        df_tot = df_tot[(df_tot["년도"] >= 2018) & (df_tot["년도"] <= 2026)]
        df_tot = df_tot.drop_duplicates(subset=["년도", "소속", "성명", "징계종류", "징계일"])
        df_tot = df_tot.sort_values(by=["년도", "징계일", "성명"]).reset_index(drop=True)
        df_tot["번호"] = df_tot.index + 1
        return df_tot[COLS + ["소속_정제", "징계종류_정제"]]
    except:
        return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])

df = load_data()

st.subheader("Filters")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    o_yr = sorted(list(df["년도"].unique()))
    o_dp = sorted(list(df["소속_정제"].unique()))
    o_nm = sorted(list(df["성명"].unique()))
    with col1: f_yr = st.multiselect("Year", options=o_yr)
    with col2: f_dept = st.multiselect("Department", options=o_dp)
    with col3: f_name = st.multiselect("Name", options=o_nm)
    
    f_df = df.copy()
    if f_yr: f_df = f_df[f_df["년도"].isin(f_yr)]
    if f_dept: f_df = f_df[f_df["소속_정제"].isin(f_dept)]
    if f_name: f_df = f_df[f_df["성명"].isin(f_name)]
else:
    f_df = df

st.markdown("---")
if not f_df.empty:
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total Count", f"{len(f_df)} 건")
    kpi2.metric("Total Depts", f"{f_df['소속_정제'].nunique()} 개소")
    kpi3.metric("Top Type", f"{f_df['징계종류_정제'].value_counts().idxmax()}")

st.subheader("Charts")
if not f_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.write("#### Company Bar Chart")
        st.plotly_chart(px.bar(f_df, x="구분", color="구분"), use_container_width=True)
        st.write("#### Department Histogram")
        st.plotly_chart(px.histogram(f_df, x="소속_정제", color="징계종류_정제", barmode="stack"), use_container_width=True)
    with c2:
        st.write("#### Yearly Trend Line")
        yt = f_df.groupby("년도").size().reset_index(name="count")
        st.plotly_chart(px.line(yt, x="년도", y="count", markers=True), use_container_width=True)
        st.write("#### Discipline Pie Chart")
        st.plotly_chart(px.pie(f_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("No data found.")

st.markdown("---")
st.subheader("Data List")
if not f_df.empty:
    disp_df = f_df[COLS]
    st.dataframe(disp_df, use_container_width=True, hide_index=True)
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: disp_df.to_excel(w, index=False)
    st.download_button(label="Download Excel", data=output.getvalue(), file_name="report.xlsx")
