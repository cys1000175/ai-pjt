import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(layout="wide")
st.title("Discipline Integrated Dashboard")

EXCEL_FILE = "data.xlsx"
COLS = ["번호", "년도", "구분", "소속", "직책", "성명", "징계 사유", "징계종류", "징계일"]

def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])
    try:
        excel_obj = pd.ExcelFile(EXCEL_FILE)
        rows = []
        for s in excel_obj.sheet_names:
            # 1단계: 첫 행부터 서식 검사 없이 무조건 다 읽어옵니다.
            df_s = excel_obj.parse(s, header=None)
            if df_s.empty: continue
            
            # 2단계: '번호'나 '년도'가 보이는 행을 컬럼명으로 지정
            hdr = 0
            for i, r in df_s.iterrows():
                v = [str(x).strip() for x in r.values if pd.notna(x)]
                if any(k in v for k in ["번호", "년도", "성명", "구분", "소속"]):
                    hdr = i
                    break
            
            df_real = excel_obj.parse(s, skiprows=hdr)
            df_real.columns = [str(c).strip() for c in df_real.columns]
            
            # 3단계: 없는 열이 있어도 에러 안 나게 유연하게 매칭
            for c in COLS:
                if c not in df_real.columns:
                    m = [x for x in df_real.columns if c in x or x in c]
                    df_real[c] = df_real[m[0]] if m else ""
            
            rows.append(df_real[COLS])
            
        if not rows: return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])
        
        # 4단계: 비어있지 않은 모든 행을 강제로 결합
        df_tot = pd.concat(rows, ignore_index=True).dropna(how="all")
        
        # 이름이나 소속이 없어도 강제로 공백 처리해서 화면에 출력되게 잠금 해제
        df_tot["성명"] = df_tot["성명"].fillna("미상").astype(str).str.strip()
        df_tot = df_tot[df_tot["성명"] != ""]
        df_tot["구분"] = df_tot["구분"].fillna("미지정").astype(str).str.strip().replace("", "미지정")
        df_tot["소속"] = df_tot["소속"].fillna("미지정").astype(str).str.strip().replace("", "미지정")
        df_tot["징계종류"] = df_tot["징계종류"].fillna("기타").astype(str).str.strip().replace("", "기타")
        
        # 정제용 임시 데이터 생성
        df_tot["소속_정제"] = df_tot["소속"].apply(lambda x: x.split("(")[0].strip() if "(" in str(x) else str(x).strip())
        df_tot["징계종류_정제"] = df_tot["징계종류"].apply(lambda x: next((t for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계"] if t in str(x)), "기타"))
        df_tot["년도"] = pd.to_numeric(df_tot["년도"], errors='coerce').fillna(2026).astype(int)
        
        df_tot = df_tot.sort_values(by=["년도", "징계일"]).reset_index(drop=True)
        df_tot["번호"] = df_tot.index + 1
        return df_tot
    except:
        return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])

if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_data()

df = st.session_state.discipline_data

st.subheader("Search Filter")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    o_yr = sorted(list(df["년도"].unique()))
    o_dp = sorted(list(df["소속_정제"].unique()))
    o_nm = sorted(list(df["성명"].unique()))
    
    with col1: f_yr = st.multiselect("Year", options=o_yr)
    with col2: f_dp = st.multiselect("Department", options=o_dp)
    with col3: f_nm = st.multiselect("Name", options=o_nm)
    
    f_df = df.copy()
    if f_yr: f_df = f_df[f_df["년도"].isin(f_yr)]
    if f_dp: f_df = f_df[f_df["소속_정제"].isin(f_dp)]
    if f_nm: f_df = f_df[f_df["성명"].isin(f_nm)]
else:
    f_df = df

st.markdown("---")
st.subheader("Visualization Chart")
if not f_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Total Records by Company")
        st.plotly_chart(px.bar(f_df, x="구분", color="구분"), use_container_width=True)
        st.markdown("#### Records by Department")
        st.plotly_chart(px.histogram(f_df, x="소속_정제", color="징계종류_정제", barmode="stack"), use_container_width=True)
