import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os
import re

# 1. 페이지 설정
st.set_page_config(page_title="통합 징계 내역 관리 시스템", layout="wide")
st.title("📊 법인별·사업장별 징계 내역 통합 대시보드 (2018~2026)")
st.markdown("AI 공모전 제출용 다중 시트 자동 통합 및 데이터 정제 시스템 프로토타입입니다.")

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
    # 괄호 내용 및 하위 부서 디테일 제거 (예: 서울본사 (미화) -> 서울본사)
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            # [🔥 다중 시트 통합 핵심] sheet_name=None 으로 설정하여 모든 시트를 딕셔너리 형태로 읽어옵니다.
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
            
            combined_rows = []
            
            # 각 시트(년도별 시트 등)를 순회하며 데이터 추출
            for sheet_name, raw_df in all_sheets.items():
                header_row_idx = None
                
                # 시트 내부에서 진짜 데이터 시작점(헤더) 찾기
                for idx, row in raw_df.iterrows():
                    row_str = [str(x) for x in row.values]
                    if any("번호" in s or "년도" in s or "구분" in s for s in row_str):
                        header_row_idx = idx
                        break
                
                # 헤더 위치를 찾았다면 해당 위치부터 시트 데이터를 바르게 파싱
                if header_row_idx is not None:
                    sheet_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, skiprows=header_row_idx)
                else:
                    sheet_df = raw_df.copy()
                
                # 열 이름 공백 정리
                sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                
                # 데이터가 없는 시트는 패스
                if sheet_df.empty:
                    continue
                    
                # 필수 열 검사 및 매핑 보정
                for col in REQUIRED_COLUMNS:
                    if col not in sheet_df.columns:
                        matched_col = [c for c in sheet_df.columns if col in c or c in col]
                        if matched_col:
                            sheet_df[col] = sheet_df[matched_col[0]]
                        else:
                            sheet_df[col] = ""
                
                # 유효한 행만 필터링하여 리스트에 축적
                sheet_df = sheet_df[REQUIRED_COLUMNS]
                combined_rows.append(sheet_df)
            
            # 모든 시트 데이터가 모였다면 하나로 합치기
            if combined_rows:
                total_df = pd.concat(combined_rows, ignore_index=True)
            else:
                return get_sample_data()
                
            # 데이터 데이터 정제 및 가공 연산 수행
            total_df = total_df.dropna(subset=["성명", "징계종류"], how="all") # 완전 빈 줄 제거
            total_df["구분"] = total_df["구분"].fillna("미지정").astype(str).str.strip()
            total_df["구분"] = total_df["구분"].apply(lambda x: "미지정" if x == "" or x == "nan" else x)
            
            # 사업장명(소속) 및 징계종류 유사어 자동 매핑 정리
            total_df["소속_정제"] = total_df["소속"].apply(clean_location_name)
            total_df["징계종류_정제"] = total_df["징계종류"].apply(clean_discipline_type)
            
            # 년도 데이터 정형화
            total_df["년도"] = pd.to_numeric(total_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            
            # 번호 순서대로 재정렬 (1부터 이쁘게 나오도록 다시 매김)
            total_df = total_df.sort_values(by=["년도", "징계일"]).reset_index(drop=True)
            total_df["번호"] = total_df.index + 1
            
            return total_df
            
        except Exception as e:
            st.error(f"❌ 다중 시트 엑셀을 통합하는 과정에서 오류가 발생했습니다: {e}")
            return get_sample_data()
    else:
        st.info("💡 'data.xlsx' 파일이 서버 컴퓨터에 존재하지 않습니다. 프로토타입용 샘플 데이터를 활성화합니다.")
        return get_sample_data()

def get_sample_data():
    return pd.DataFrame([
        {
            "번호": 1, "년도": 2026, "구분": "샘플법인", "소속": "서울본사 (미화)", "소속_정제": "서울본사",
            "직책": "대리", "성명": "홍길동", "징계 사유": "엑셀 다중시트 연동 대기 중 - 샘플 데이터입니다.", 
            "징계종류": "감봉(10%)", "징계종류_정제": "감봉", "징계일": "2026-03-15"
        }
    ])

# 세션 상태 데이터 초기화 및 보존
if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

df = st.session_state.discipline_data

# ----------------------------------------------------------------💡 사이드바: 데이터 추가 입력
st.sidebar.header("➕ 신규 징계 내역 입력")
with st.sidebar.form(key="input_form", clear_on_submit=True):
    division_options = list(df["구분"].dropna().unique()) if not df.empty else ["기본 법인"]
    if "기타" not in division_options: division_options.append("기타")
    input_division = st.selectbox("구분(법인) 선택", division_options)
    
    current_year = datetime.date.today().year
    input_year = st.number_input("년도", min_value=2018, max_value=2030, value=current_year)
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
            "번호": next_id, "년도": int(input_year), "구분": input_division, "소속": input_dept,
            "소속_정제": clean_location_name(input_dept), "직책": input_position, "성명": input_name,
            "징계 사유": input_reason, "징계종류": input_type, "징계종류_정제": clean_discipline_type(input_type),
            "징계일": input_date.strftime("%Y-%m-%d")
        }
        
        st.session_state.discipline_data = pd.concat([st.session_state.discipline_data, pd.DataFrame([new_row])], ignore_index=True)
        st.sidebar.success("✅ 대시보드에 반영되었습니다!")
        st.rerun()
    else:
        st.sidebar.error("⚠️ 필수 항목(소속, 성명, 징계 사유)을 입력해주세요.")

df = st.session_state.discipline_data

# ----------------------------------------------------------------🔍 메인 화면: 동적 필터
st.subheader("🔍 데이터 통합 필터링 (2
