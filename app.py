import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os
import re

# 1. 페이지 설정
st.set_page_config(page_title="징계 내역 관리 시스템", layout="wide")
st.title("📊 법인별·사업장별 징계 내역 관리 대시보드")
st.markdown("AI 공모전 제출용 통합 관리 및 데이터 자동 정제 시스템 프로토타입입니다.")

# 2. 엑셀 데이터 로드 및 자동 정제(텍스트 표준화) 설정
EXCEL_FILE = "data.xlsx"
REQUIRED_COLUMNS = ["번호", "년도", "구분", "소속", "직책", "성명", "징계 사유", "징계종류", "징계일"]

def clean_discipline_type(val):
    if pd.isna(val) or not isinstance(val, str):
        return "기타"
    val = val.strip()
    if "해고" in val: return "해고"
    elif "강등" in val: return "강등"
    elif "정직" in val: return "정직"
    elif "감봉" in val: return "감봉"
    elif "견책" in val: return "견책"
    elif "경고" in val: return "경고"
    elif "훈계" in val: return "훈계"
    elif "권고" in val or "사직" in val: return "권고사직"
    return "기타"

def clean_location_name(val):
    if pd.isna(val) or not isinstance(val, str):
        return "기타 사업장"
    val = val.replace('\n', ' ').strip()
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            read_df = pd.read_excel(EXCEL_FILE)
            
            if read_df.empty:
                st.warning(f"⚠️ '{EXCEL_FILE}' 파일이 비어 있습니다. 샘플 데이터를 표시합니다.")
                return get_sample_data()
                
            for col in REQUIRED_COLUMNS:
                if col not in read_df.columns:
                    matched_col = [c for c in read_df.columns if c.strip() == col]
                    if matched_col:
                        read_df[col] = read_df[matched_col[0]]
                    else:
                        read_df[col] = ""
            
            read_df["구분"] = read_df["구분"].fillna("미지정").astype(str).str.strip()
            read_df["구분"] = read_df["구분"].apply(lambda x: "미지정" if x == "" or x == "nan" else x)
            
            read_df["소속_정제"] = read_df["소속"].apply(clean_location_name)
            read_df["징계종류_정제"] = read_df["징계종류"].apply(clean_discipline_type)
            
            read_df["년도"] = pd.to_numeric(read_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            return read_df
            
        except Exception as e:
            st.error(f"❌ 엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
            return get_sample_data()
    else:
        st.error(f"🚨 서버 컴퓨터(GitHub)에서 '{EXCEL_FILE}' 파일을 찾을 수 없습니다! 파일 이름을 정확히 소문자 data.xlsx 로 해서 업로드했는지 확인해 주세요. 임시 샘플 데이터를 노출합니다.")
        return get_sample_data()

def get_sample_data():
    return pd.DataFrame([
        {
            "번호": 1, "년도": 2026, "구분": "샘플법인(A사)", "소속": "서울본사 (미화)", "소속_정제": "서울본사",
            "직책": "대리", "성명": "홍길동", "징계 사유": "엑셀 연동 대기 중 - 샘플 데이터입니다.", 
            "징계종류": "감봉(10%)", "징계종류_정제": "감봉", "징계일": "2026-03-15"
        }
    ])

# 세션 상태에 데이터 저장
if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

df = st.session_state.discipline_data

# ----------------------------------------------------------------💡 사이드바: 데이터 추가 입력
st.sidebar.header("➕ 신규 징계 내역 입력")
with st.sidebar.form(key="input_form", clear_on_submit=True):
    division_options = list(df["구분"].dropna().unique()) if not df.empty else ["A전자", "B화학", "C건설"]
    if "기타" not in division_options: division_options.append("기타")
    input_division = st.selectbox("구분(법인) 선택", division_options)
    
    current_year = datetime.date.today().year
    input_year = st.number_input("년도", min_value=2020, max_value=2030, value=current_year)
    input_date = st.date_input("징계일 선택", datetime.date.today())
    
    input_dept = st.text
