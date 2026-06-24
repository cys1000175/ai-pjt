import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(layout="wide")
st.title("🏛️ 인사 징계 내역 통합 관리 시스템")
st.markdown("##### Enterprise HR Data Platform")

ADMIN_PASSWORD = "1234"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.write("### 🔒 보안 잠금 시스템")
    st.info("비밀번호를 입력해 주세요.")
    input_pw = st.text_input("비밀번호", type="password")
    if st.button("인증하기"):
        if input_pw == ADMIN_PASSWORD:
            st.session_state.authenticated = True
            st.success("🔒 인증 성공!")
            st.rerun()
        else:
            st.error("❌ 비밀번호 불일치")
    st.stop()

FILE_NAME = "data.xlsx"

@st.cache_data
def load_all_data():
    if not os.path.exists(FILE_NAME): 
        return pd.DataFrame()
    try:
        excel = pd.ExcelFile(FILE_NAME)
        all_sheets = []
        for sheet in excel.sheet_names:
            df = excel.parse(sheet, header=None)
            if df.empty: continue
            start_idx = 0
            for idx, row in df.iterrows():
                row_str = [str(x) for x in row.values if pd.notna(x)]
                if any("번호" in s for s in row_str):
                    start_idx = idx
                    break
            data_df = excel.parse(sheet, skiprows=start_idx)
            data_df.columns = [str(c).strip() for c in data_df.columns]
            name_col = [c for c in data_df.columns if "성명" in c or "성 명" in c]
            if name_col: 
                data_df = data_df.dropna(subset=[name_col[0]])
            else: 
                continue
            res = pd.DataFrame()
            def get_val(keywords, default=""):
                for c in data_df.columns:
                    if any(k in c for k in keywords): 
                        return data_df[c].fillna(default)
                return default
            res["Year"] = get_val(["년도", "연도"], sheet.replace("년", "").strip())
            res["Division"] = get_val(["구분"])
            res["Date"] = get_val(["일 자", "일자", "징계일"])
            res["Dept"] = get_val(["소속"])
            res["Position"] = get_val(["직책", "직급"])
            res["Name"] = get_val(["성명", "성 명"])
            res["Reason"] = get_val(["사유", "징계 사유"])
            res["Type"] = get_val(["종류", "징계 종류"])
            all_sheets.append(res)
        if not all_sheets: 
            return pd.DataFrame()
        final_df = pd.concat(all_sheets, ignore_index=True)
        final_df["Name"] = final_df["Name"].astype(str).str.strip()
        final_df = final_df[final_df["Name"] != ""]
        final_df["Dept_Clean"] = final_df["Dept"].astype(str).apply(lambda x: x.split("\n")[0].split("(")[0].strip())
        final_df["Type_Clean"] = final_df["Type"].astype(str).apply(lambda x: next((t for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고"] if t in x), "기타"))
        def backfill_division(row):
            div = str(row["Division"]).strip()
            if div in ["", "nan", "None"] or pd.isna(row["Division"]):
                r = str(row["Reason"])
                if any(k in r for k in ["괴롭힘", "성희롱", "배임"]): 
                    return "인사위 결정"
                return "일반 징계위원회"
            return div
        final_df["Division"] = final_df.apply(backfill_division, axis=1)
        final_df["Year"] = pd.to_numeric(final_df["Year"], errors="coerce").fillna(2026).astype(int)
        final_df = final_df[(final_df["Year"] >= 2018) & (final_df["Year"] <= 2026)]
        return final_df.sort_values(by=["Year", "Date"]).reset_index(drop=True)
    except: 
        return pd.DataFrame()

df = load_all_data()

if df.empty:
    st.error("데이터 로드 실패")
else:
    st.markdown("### 🔍 필터 컨트롤 타워")
    c1, c2, c3 = st.columns(3)
    with c1: f_yr = st.multiselect("연도 선택", options=sorted(df["Year"].unique()))
    with c2: f_dp = st.multiselect("소속 사업장 선택", options=sorted(df["Dept_Clean"].unique()))
    with c3: f_nm = st.multiselect("성명 검색", options=sorted(df["Name"].unique()))
        
    f_df = df.copy()
    if f_yr: f_df = f_df[f_df["Year"].isin(f_yr)]
    if f_dp: f_df = f_df[f_df["Dept_Clean"].isin(f_dp)]
    if f_nm: f_df = f_df[f_df["Name"].
