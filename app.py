import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re

st.set_page_config(page_title="HR Dashboard", layout="wide")
st.title("📊 징계 내역 통합 관리 대시보드 (2018~2026)")
st.markdown("AI 공모전 제출용 전 시트 자동 파싱 및 데이터 표준화 엔지니어링 시스템입니다.")

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
    val = val.replace('\n', ' ').strip()
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip()

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
            
            # 2026년 시트 등의 우측 통계 집계 열 노이즈 제거
            df_real = df_real.iloc[:, :11] 
            
            for c in COLS:
                if c not in df_real.columns:
                    m = [x for x in df_real.columns if c in x or x in c]
                    df_real[c] = df_real[m[0]] if m else ""
            
            # 일 자 열이 존재할 경우 징계일로 대 대체 연동
            if "징계일" not in df_real.columns or df_real["징계일"].dropna().empty:
                if "일 자" in df_real.columns:
                    df_real["징계일"] = df_real["일 자"]
            
            rows.append(df_real[COLS])
            
        if not rows: return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])
        
        df_tot = pd.concat(rows, ignore_index=True)
        df_tot = df_tot.dropna(subset=["성명", "소속", "징계종류"], how="all")
        
        # 데이터 정제 전처리 패키지
        df_tot["성명"] = df_tot["성명"].fillna("미상").astype(str).str.strip()
        df_tot = df_tot[df_tot["성명"] != ""]
        df_tot["구분"] = df_tot["구분"].fillna("인사위").astype(str).str.strip().replace("", "인사위")
        df_tot["소속"] = df_tot["소속"].fillna("미지정").astype(str).str.strip()
        df_tot["징계종류"] = df_tot["징계종류"].fillna("기타").astype(str).str.strip()
        
        df_tot["소속_정제"] = df_tot["소속"].apply(clean_loc)
        df_tot["징계종류_정제"] = df_tot["징계종류"].apply(clean_type)
        
        # 날짜 기반 년도 자동 보정 엔진
        df_tot["징계일"] = df_tot["징계일"].astype(str).str.split(" ").str[0]
        df_tot["년도_추출"] = df_tot["징계일"].str.split("-").str[0]
        df_tot["년도"] = pd.to_numeric(df_tot["년도_추출"], errors='coerce').fillna(df_tot["년도"]).fillna(2026)
        df_tot["년도"] = pd.to_numeric(df_tot["년도"], errors='coerce').fillna(2026).astype(int)
        
        # 유효 범위 스크리닝 (2018~2026)
        df_tot = df_tot[(df_tot["년도"] >= 2018) & (df_tot["년도"] <= 2026)]
        
        # 중복 데이터 제거 처리 및 인덱싱 재정렬
        df_tot = df_tot.drop_duplicates(subset=["년도", "소속", "성명", "징계종류", "징계일"])
        df_tot = df_tot.sort_values(by=["년도", "징계일", "성명"]).reset_index(drop=True)
        df_tot["번호"] = df_tot.index + 1
        
        return df_tot[COLS + ["소속_정제", "징계종류_정제"]]
    except:
        return pd.DataFrame(columns=COLS + ["소속_정제", "징계종류_정제"])

df = load_data()

# ---------------------------------------------------------------- 검색 필터 영역
st.subheader("🔍 데이터 통합 다차원 검색 필터")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    o_yr = sorted(list(df["년도"].unique()))
    o_dp = sorted(list(df["소속_정제"].unique()))
    o_nm = sorted(list(df["성명"].unique()))
    
    with col1: f_yr = st.multiselect("연도 선택 (비워두면 전체)", options=o_yr)
    with col2: f_dp = st.multiselect("소속 사업장 선택 (비워두면 전체)", options=o_dp)
    with col3: f_name = st.multiselect("성명 검색 (비워두면 전체)", options=o_nm)
    
    f_df = df.copy()
    if f_yr: f_df = f_df[f_df["년도"].isin(f_yr)]
    if f_dp: f_df = f_df[f_df["소속_정제"].isin(f_dp)]
    if f_name: f_df = f_df[f_df["성명"].isin(f_name)]
else:
    f_df = df

# ---------------------------------------------------------------- 핵심 KPI 지표 컴포넌트
st.markdown("---")
if not f_df.empty:
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("📊 총 누적 발생 건수", f"{len(f_df)} 건")
    kpi2.metric("🏢 분석 대상 사업장 수", f"{f_df['소속_정제'].nunique()} 개소")
    kpi3.metric("⚖️ 가장 빈번한 징계 종류", f"{f_df['징계종류_정제'].value_counts().idxmax()}")

# ---------------------------------------------------------------- 시각화 대시보드 영역
st.subheader("📈 통계 분석 및 데이터 시각화")
if not f_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📍 소속 사업장별 징계 발생 TOP 10")
        top_10_loc = f_df["소속_정제"].value_counts().head(10).reset_index()
        top_10_loc.columns = ["소속_정제", "건수"]
        st.plotly_chart(px.bar(top_10_loc, x="소속_정제", y="건수", color="소속_정제", labels={"소속_정제": "사업장명"}), use_container_width=True)
        
        st.markdown("#### 🗂️ 사업장별 징계 종류 적층 분포")
        st.plotly_chart(px.histogram(f_df, x="소속_정제", color="징계종류_정제", barmode="stack", labels={"소속_정제": "사업장명", "징계종류_정제": "징계종류"}), use_container_width=True)
    with c2:
        st.markdown("#### 📅 연도별 장기 추이 트렌드 (2018-2026)")
        yt = f_df.groupby("년도").size().reset_index(name="발생건수")
        st.plotly_chart(px.line(yt, x="년도", y="발생건수", text="발생건수", markers=True), use_container_width=True)
        
        st.markdown("#### ⚖️ 전체 징계 유형별 비율 현황")
        st.plotly_chart(px.pie(f_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("선택 조건에 부합하는 분석 데이터가 없습니다.")

# ---------------------------------------------------------------- 상세 데이터 리스트 테이블 영역
st.markdown("---")
st.subheader("📋 상세 내역 마스터 테이블")
if not f_df.empty:
    disp_df = f_df[COLS]
    st.dataframe(disp_df, use_container_width=True, hide_index=True)
    
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: 
        disp_df.to_excel(w, index=False)
    st.download_button(label="📥 통합 정제 데이터 다운로드 (Excel)", data=output.getvalue(), file_name="HR_integrated_report.xlsx")
else:
    st.info("조회된 데이터가 없습니다. data.xlsx 파일을 다시 점검해 주세요.")
