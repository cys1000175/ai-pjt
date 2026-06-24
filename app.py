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
    """징계종류 텍스트 유연하게 매핑 (빈 값이면 '기타' 처리하여 필터 깨짐 방지)"""
    if pd.isna(val) or not isinstance(val, str) or val.strip() == "":
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
    elif "전배" in val: return "전배"
    return "기타"

def clean_location_name(val):
    """소속 사업장명 뒤의 줄바꿈이나 부서 괄호 요소를 제거하여 자동 통합"""
    if pd.isna(val) or not isinstance(val, str) or val.strip() == "":
        return "미지정 소속"
    val = val.replace('\n', ' ').strip()
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip() if val.strip() != "" else "미지정 소속"

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            read_df = pd.read_excel(EXCEL_FILE)
            
            # 1. 존재하지 않는 열이 있다면 공백으로 자동 생성
            for col in REQUIRED_COLUMNS:
                if col not in read_df.columns:
                    read_df[col] = ""
            
            # 2. 모든 텍스트 컬럼의 양끝 공백 제거 및 결측치(NaN) 방어 처리
            for col in REQUIRED_COLUMNS:
                if col not in ["번호", "년도"]:
                    read_df[col] = read_df[col].fillna("").astype(str).str.strip()
            
            # 3. 구분(법인) 빈 값 자동 메움
            read_df["구분"] = read_df["구분"].apply(lambda x: "미지정 법인" if x == "" else x)
            
            # 4. 텍스트 유사성 기반 자동 정제 열 매핑
            read_df["소속_정제"] = read_df["소속"].apply(clean_location_name)
            read_df["징계종류_정제"] = read_df["징계종류"].apply(clean_discipline_type)
            
            # 5. 년도 숫자 변환 예외 처리
            read_df["년도"] = pd.to_numeric(read_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            
            return read_df
        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
            return pd.DataFrame(columns=REQUIRED_COLUMNS + ["소속_정제", "징계종류_정제"])
    else:
        # 파일이 없을 때 샘플 데이터 자동 작동
        return pd.DataFrame([
            {
                "번호": 1, "년도": 2026, "구분": "A전자", "소속": "서울본사 (미화)", "소속_정제": "서울본사",
                "직책": "대리", "성명": "홍길동", "징계 사유": "근태 불량", 
                "징계종류": "감봉(10%)", "징계종류_정제": "감봉", "징계일": "2026-03-15"
            }
        ])

# 세션 상태 데이터 로드
if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

df = st.session_state.discipline_data

# ----------------------------------------------------------------💡 사이드바: 데이터 추가 입력
st.sidebar.header("➕ 신규 징계 내역 입력")
with st.sidebar.form(key="input_form", clear_on_submit=True):
    division_options = list(df["구분"].unique()) if not df.empty else ["A전자", "B화학", "C건설"]
    input_division = st.selectbox("구분(법인) 선택", division_options)
    
    current_year = datetime.date.today().year
    input_year = st.number_input("년도", min_value=2020, max_value=2030, value=current_year)
    input_date = st.date_input("징계일 선택", datetime.date.today())
    
    input_dept = st.text_input("소속(사업장명)")
    input_position = st.text_input("직책")
    input_name = st.text_input("성명")
    input_type = st.selectbox("징계종류", ["견책", "감봉", "정직", "강등", "해고", "경고", "훈계", "권고사직"])
    input_reason = st.text_area("징계 사유")
    
    submit_button = st.form_submit_button(label="대시보드에 추가")

if submit_button:
    if input_dept and input_name and input_reason:
        next_id = len(st.session_state.discipline_data) + 1
        new_row = {
            "번호": next_id, "년도": int(input_year), "구분": input_division,
