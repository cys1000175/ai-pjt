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
st.subheader("🔍 데이터 필터링 (유사 내용 자동 정리 반영)")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    with col1:
        # 데이터 유실 방지를 위해 빈 값 옵션 제거 처리 후 셀렉트박스 구축
        options_div = [x for x in df["구분"].unique() if x != ""]
        f_division = st.multiselect("구분(법인)", options=options_div, default=options_div)
    with col2:
        options_yr = sorted(list(df["년도"].unique()))
        f_year = st.multiselect("년도", options=options_yr, default=options_yr)
    with col3:
        options_dept = [x for x in df["소속_정제"].unique() if x != ""]
        f_dept = st.multiselect("소속(자동 정리됨)", options=options_dept, default=options_dept)

    # 필터 조건 적용 (조건에 해당하지 않더라도 데이터가 아예 증발하지 않도록 안전 매핑)
    filtered_df = df[
        (df["구분"].isin(f_division)) & 
        (df["년도"].isin(f_year)) & 
        (df["소속_정제"].isin(f_dept))
    ]
else:
    filtered_df = df

# ----------------------------------------------------------------📊 메인 화면: 시각화 차트
st.markdown("---")
st.subheader("📊 실시간 분석 통계")

if not filtered_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🏢 구분(법인)별 발생 건수")
        st.plotly_chart(px.bar(filtered_df, x="구분", color="구분", labels={"구분": "법인 구분", "count": "건수"}), use_container_width=True)
        
        st.markdown("#### 📍 소속(사업장)별/징계종류 분포 (통합 통계)")
        st.plotly_chart(px.histogram(filtered_df, x="소속_정제", color="징계종류_정제", barmode="stack", labels={"소속_정제": "소속 사업장", "징계종류_정제": "징계 종류"}), use_container_width=True)
    with c2:
        st.markdown("#### 📅 월별 트렌드 (징계일 기준)")
        trend = filtered_df.copy()
        # 다양한 날짜 텍스트 형식 강제 연산 처리
        trend["징계일"] = pd.to_datetime(trend["징계일"], errors='coerce')
        trend = trend.dropna(subset=["징계일"])
        
        if not trend.empty:
            trend["년월"] = trend["징계일"].dt.strftime("%Y-%m")
            trend_g = trend.groupby("년월").size().reset_index(name="건수").sort_values("년월")
            st.plotly_chart(px.line(trend_g, x="년월", y="건수", markers=True), use_container_width=True)
        else:
            st.info("시계열 그래프용 '징계일'이 올바른 날짜 형식(예: 2026-03-15)이 아닙니다. 아래 데이터 표를 확인하세요.")
        
        st.markdown("#### ⚖️ 징계종류 비율 (통합 통계)")
        st.plotly_chart(px.pie(filtered_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("⚠️ 필터박스에서 모든 체크를 해제하셨거나 일치하는 데이터가 없습니다. 상단 필터에서 항목을 선택해 주세요.")

# ----------------------------------------------------------------📄 메인 화면: 데이터 표 및 다운로드
st.markdown("---")
st.subheader("📋 상세 내역 리스트")
display_df = filtered_df[REQUIRED_COLUMNS] if not filtered_df.empty else df[REQUIRED_COLUMNS]
st.dataframe(display_df, use_container_width=True, hide_index=True)

# 다운로드 버튼
import io
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    display_df.to_excel(writer, index=False)
st.download_button(
    label="📥 현재 필터링된 데이터 엑셀 다운로드", data=output.getvalue(),
    file_name="discipline_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
