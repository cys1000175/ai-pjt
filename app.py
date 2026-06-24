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
    """징계종류 텍스트 유연하게 매핑 (예: 감봉(10%) -> 감봉)"""
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
    elif "전배" in val: return "전배"
    return "기타"

def clean_location_name(val):
    """소속 사업장명 뒤의 줄바꿈이나 (미화), (보안) 등의 괄호 요소를 제거하여 자동 통합"""
    if pd.isna(val) or not isinstance(val, str):
        return "기타 사업장"
    # 줄바꿈 제거 및 괄호 내용 제거
    val = val.replace('\n', ' ').strip()
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            read_df = pd.read_excel(EXCEL_FILE)
            
            # 엑셀에 필수 열들이 모두 있는지 확인하고 없으면 빈 열 생성
            for col in REQUIRED_COLUMNS:
                if col not in read_df.columns:
                    read_df[col] = ""
            
            # 원본 데이터 백업 보존하면서 정제 데이터 열 추가 생성 및 값 변경
            read_df["구분"] = read_df["구분"].fillna("미지정").astype(str).str.strip()
            read_df["구분"] = read_df["구분"].apply(lambda x: "미지정" if x == "" else x)
            
            # [자동 정리 핵심 부문] 값 유사성 매핑 가공
            read_df["소속_정제"] = read_df["소속"].apply(clean_location_name)
            read_df["징계종류_정제"] = read_df["징계종류"].apply(clean_discipline_type)
            
            # 데이터 타입 정제
            read_df["년도"] = pd.to_numeric(read_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            return read_df
        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
            return pd.DataFrame(columns=REQUIRED_COLUMNS + ["소속_정제", "징계종류_정제"])
    else:
        # 파일이 없을 때를 대비한 기본 샘플 데이터 구조
        return pd.DataFrame([
            {
                "번호": 1, "년도": 2026, "구분": "A전자", "소속": "서울본사 (미화)", "소속_정제": "서울본사",
                "직책": "대리", "성명": "홍길동", "징계 사유": "근태 불량", 
                "징계종류": "감봉(10%)", "징계종류_정제": "감봉", "징계일": "2026-03-15"
            }
        ])

# 세션 상태에 데이터 저장
if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

# 데이터 참조 규칙 정의
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
    
    input_dept = st.text_input("소속(사업장명)")
    input_position = st.text_input("직책")
    input_name = st.text_input("성명")
    input_type = st.selectbox("징계종류", ["견책", "감봉", "정직", "강등", "해고", "경고", "훈계", "권고사직"])
    input_reason = st.text_area("징계 사유")
    
    submit_button = st.form_submit_button(label="대시보드에 추가")

if submit_button:
    if input_dept and input_name and input_reason:
        next_id = len(st.session_state.discipline_data) + 1
        
        # [오류 해결] 괄호 쌍이 꼬여서 발생했던 SyntaxError 부분을 깔끔하게 정돈했습니다.
        new_row = {
            "번호": next_id,
            "년도": int(input_year),
            "구분": input_division,
            "소속": input_dept,
            "소속_정제": clean_location_name(input_dept),
            "직책": input_position,
            "성명": input_name,
            "징계 사유": input_reason,
            "징계종류": input_type,
            "징계종류_정제": clean_discipline_type(input_type),
            "징계일": input_date.strftime("%Y-%m-%d")
        }
        
        st.session_state.discipline_data = pd.concat([st.session_state.discipline_data, pd.DataFrame([new_row])], ignore_index=True)
        st.sidebar.success("✅ 대시보드에 반영되었습니다!")
        st.rerun()
    else:
        st.sidebar.error("⚠️ 필수 항목(소속, 성명, 징계 사유)을 입력해주세요.")

df = st.session_state.discipline_data

# ----------------------------------------------------------------🔍 메인 화면: 동적 필터 (정제된 텍스트 기준)
st.subheader("🔍 데이터 필터링 (유사 내용 자동 정리 반영)")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    with col1:
        f_division = st.multiselect("구분(법인)", options=df["구분"].unique(), default=df["구분"].unique())
    with col2:
        f_year = st.multiselect("년도", options=sorted(df["년도"].unique()), default=sorted(df["년도"].unique()))
    with col3:
        f_dept = st.multiselect("소속(자동 정리됨)", options=df["소속_정제"].unique(), default=df["소속_정제"].unique())

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
        trend["징계일"] = pd.to_datetime(trend["징계일"], errors='coerce')
        trend = trend.dropna(subset=["징계일"])
        
        if not trend.empty:
            trend["년월"] = trend["징계일"].dt.strftime("%Y-%m")
            trend_g = trend.groupby("년월").size().reset_index(name="건수").sort_values("년월")
            st.plotly_chart(px.line(trend_g, x="년월", y="건수", markers=True), use_container_width=True)
        else:
            st.info("시계열 그래프를 표시할 수 있는 유효한 '징계일' 데이터가 없거나 형식이 다릅니다.")
        
        st.markdown("#### ⚖️ 징계종류 비율 (통합 통계)")
        st.plotly_chart(px.pie(filtered_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("조건에 맞는 데이터가 없습니다.")

# ----------------------------------------------------------------📄 메인 화면: 데이터 표 및 다운로드
st.markdown("---")
st.subheader("📋 상세 내역 리스트")
display_df = filtered_df[REQUIRED_COLUMNS]
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
