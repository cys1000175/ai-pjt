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
    # 괄호 내용 제거 (예: 서울본사 (미화) -> 서울본사)
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            # [🔥 개선 핵심] 서식이 섞여 있는 원본 엑셀 파일을 고려하여 상단 행을 유연하게 탐색합니다.
            # '번호'나 '년도'가 포함된 행을 진짜 헤더(Header)로 잡기 위해 먼저 서식을 검사합니다.
            raw_df = pd.read_excel(EXCEL_FILE, header=None)
            header_row_idx = 0
            
            # '번호' 또는 '년도'라는 글자가 들어있는 행을 헤더 시작 위치로 자동 감지
            for idx, row in raw_df.iterrows():
                row_str = [str(x) for x in row.values]
                if any("번호" in s or "년도" in s for s in row_str):
                    header_row_idx = idx
                    break
            
            # 올바른 헤더 위치부터 데이터를 다시 로드합니다.
            read_df = pd.read_excel(EXCEL_FILE, skiprows=header_row_idx)
            
            # 열 이름 양끝 공백 제거 정리
            read_df.columns = [str(c).strip() for c in read_df.columns]
            
            # 파일은 존재하지만 유효 데이터 행이 없을 때의 방어 로직
            if read_df.empty:
                return get_sample_data()
                
            # 필수 열 검사 및 보정 기법 적용
            for col in REQUIRED_COLUMNS:
                if col not in read_df.columns:
                    # 유사한 이름의 열이 있는지 재확인 (예: '년도 ' 또는 '년度')
                    matched_col = [c for c in read_df.columns if col in c or c in col]
                    if matched_col:
                        read_df[col] = read_df[matched_col[0]]
                    else:
                        read_df[col] = ""
            
            # 필터 전처리 가공
            read_df["구분"] = read_df["구분"].fillna("미지정").astype(str).str.strip()
            read_df["구분"] = read_df["구분"].apply(lambda x: "미지정" if x == "" or x == "nan" else x)
            
            read_df["소속_정제"] = read_df["소속"].apply(clean_location_name)
            read_df["징계종류_정제"] = read_df["징계종류"].apply(clean_discipline_type)
            
            # 년도 전처리 및 결측치 제거
            read_df["년도"] = pd.to_numeric(read_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            
            return read_df
            
        except Exception as e:
            st.error(f"❌ 엑셀 파일을 해석하는 과정에서 오류가 발생했습니다: {e}")
            return get_sample_data()
    else:
        st.info("💡 'data.xlsx' 파일이 서버 컴퓨터에 존재하지 않습니다. 프로토타입용 샘플 데이터를 활성화합니다.")
        return get_sample_data()

def get_sample_data():
    return pd.DataFrame([
        {
            "번호": 1, "년도": 2026, "구분": "기본 법인", "소속": "서울본사 (미화)", "소속_정제": "서울본사",
            "직책": "대리", "성명": "홍길동", "징계 사유": "엑셀 연동 대기 중 - 시스템 정상 작동 확인용 샘플 데이터입니다.", 
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

# 새로고침 후 데이터 동기화
df = st.session_state.discipline_data

# ----------------------------------------------------------------🔍 메인 화면: 동적 필터
st.subheader("🔍 데이터 필터링 (유사 내용 자동 정리 반영)")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    with col1:
        f_division = st.multiselect("구분(법인)", options=list(df["구분"].unique()), default=list(df["구분"].unique()))
    with col2:
        f_year = st.multiselect("년도", options=sorted(list(df["년도"].unique())), default=sorted(list(df["년도"].unique())))
    with col3:
        f_dept = st.multiselect("소속(자동 정리됨)", options=list(df["소속_정제"].unique()), default=list(df["소속_정제"].unique()))

    # 필터가 비어 있지 않을 때만 안전하게 필터링 쿼리 적용
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
    st.warning("⚠️ 현재 선택된 필터 조건에 부합하는 데이터가 없습니다. 필터 선택 항목을 조정해 주세요.")

# ----------------------------------------------------------------📄 메인 화면: 데이터 표 및 다운로드
st.markdown("---")
st.subheader("📋 상세 내역 리스트")

# 출력용 원본 열 추출 및 화면 출력 보장
display_cols = [c for c in REQUIRED_COLUMNS if c in filtered_df.columns]
if filtered_df.empty:
    st.info("데이터가 존재하지 않습니다.")
else:
    display_df = filtered_df[display_cols]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        display_df.to_excel(writer, index=False)
    st.download_button(
        label="📥 현재 필터링된 데이터 엑셀 다운로드", data=output.getvalue(),
        file_name="discipline_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
