import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os

# 1. 페이지 설정
st.set_page_config(page_title="징계 내역 관리 시스템", layout="wide")
st.title("📊 법인별·사업장별 징계 내역 관리 대시보드")
st.markdown("AI 공모전 제출용 통합 관리 프로토타입 시스템입니다.")

# 2. 엑셀 데이터 로드 설정
EXCEL_FILE = "data.xlsx"

# [수정 포인트 1] @st.cache_data를 사용하면 사용자가 사이드바에서 데이터를 추가해도 화면에 반영되지 않는 문제가 생깁니다.
# 실시간 입력 데이터 반영을 위해 캐싱을 제거하거나 세션 상태와 결합해야 합니다.
def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            read_df = pd.read_excel(EXCEL_FILE)
            # 엑셀 데이터의 공백이나 데이터 타입을 깔끔하게 맞춰줍니다.
            read_df["년도"] = read_df["년도"].astype(int)
            read_df["월"] = read_df["월"].astype(int)
            return read_df
        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
            return pd.DataFrame(columns=["법인", "년도", "월", "사업장", "징계대상자", "징계유형", "사유"])
    else:
        # 파일이 없을 때를 대비한 기본 샘플 데이터 구조
        return pd.DataFrame([
            {"법인": "A전자", "년도": 2026, "월": 1, "사업장": "서울본사", "징계대상자": "홍길동", "징계유형": "감봉", "사유": "근태 불량"}
        ])

# 세션 상태에 데이터 저장
if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

# [수정 포인트 2] 데이터 동기화 오류 방지를 위해 항상 세션 데이터를 직접 참조합니다.
df = st.session_state.discipline_data

# ----------------------------------------------------------------💡 사이드바: 데이터 추가 입력
st.sidebar.header("➕ 신규 징계 내역 입력")
with st.sidebar.form(key="input_form", clear_on_submit=True):
    company_options = list(df["법인"].unique()) if not df.empty else ["A전자", "B화학", "C건설"]
    if "기타" not in company_options: company_options.append("기타")
    
    input_company = st.selectbox("법인 선택", company_options)
    current_year = datetime.date.today().year
    input_year = st.number_input("년도", min_value=2020, max_value=2030, value=current_year)
    input_month = st.slider("월", min_value=1, max_value=12, value=datetime.date.today().month)
    
    input_site = st.text_input("사업장명")
    input_name = st.text_input("징계 대상자 성명")
    input_type = st.selectbox("징계 유형", ["견책", "감봉", "정직", "강등", "해고"])
    input_reason = st.text_area("징계 사유 요약")
    
    submit_button = st.form_submit_button(label="대시보드에 추가")

if submit_button:
    if input_site and input_name and input_reason:
        new_row = {
            "법인": input_company, "년도": int(input_year), "월": int(input_month),
            "사업장": input_site, "징계대상자": input_name, "징계유형": input_type, "사유": input_reason
        }
        # 데이터 합치기 및 세션 갱신
        st.session_state.discipline_data = pd.concat([st.session_state.discipline_data, pd.DataFrame([new_row])], ignore_index=True)
        st.sidebar.success("✅ 대시보드에 반영되었습니다!")
        st.rerun()
    else:
        st.sidebar.error("⚠️ 모든 항목을 입력해주세요.")

# 데이터프레임 최신화
df = st.session_state.discipline_data

# ----------------------------------------------------------------🔍 메인 화면: 동적 필터
st.subheader("🔍 데이터 필터링")
if not df.empty:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        f_company = st.multiselect("법인", options=df["법인"].unique(), default=df["법인"].unique())
    with col2:
        f_year = st.multiselect("년도", options=sorted(df["년도"].unique()), default=sorted(df["년도"].unique()))
    with col3:
        f_month = st.multiselect("월", options=sorted(df["월"].unique()), default=sorted(df["월"].unique()))
    with col4:
        f_site = st.multiselect("사업장", options=df["사업장"].unique(), default=df["사업장"].unique())

    filtered_df = df[
        (df["법인"].isin(f_company)) & (df["년도"].isin(f_year)) &
        (df["월"].isin(f_month)) & (df["사업장"].isin(f_site))
    ]
else:
    filtered_df = df

# ----------------------------------------------------------------📊 메인 화면: 시각화 차트
st.markdown("---")
st.subheader("📊 실시간 분석 통계")

if not filtered_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🏢 법인별 발생 건수")
        st.plotly_chart(px.bar(filtered_df, x="법인", color="법인"), use_container_width=True)
        st.markdown("#### 📍 사업장별/유형별 분포")
        st.plotly_chart(px.histogram(filtered_df, x="사업장", color="징계유형", barmode="stack"), use_container_width=True)
    with c2:
        st.markdown("#### 📅 월별 트렌드")
        trend = filtered_df.copy()
        
        # [수정 포인트 3 - 핵심 오류 해결] 
        # 기존 코드인 trend["월"].astype(str).str.zfill(2)는 판다스 버전에 따라 에러를 뱉습니다.
        # 안전하게 수치형 데이터를 문자로 바꾸고 가공하는 정석적인 방식으로 변경했습니다.
        trend["년월"] = trend["년도"].astype(str) + "-" + trend["월"].apply(lambda x: str(int(x)).zfill(2))
        
        # 년월 기준으로 묶어 정렬 후 차트를 그립니다.
        trend = trend.groupby("년월").size().reset_index(name="건수").sort_values("년월")
        st.plotly_chart(px.line(trend, x="년월", y="건수", markers=True), use_container_width=True)
        
        st.markdown("#### ⚖️ 징계 유형 비율")
        st.plotly_chart(px.pie(filtered_df, names="징계유형", hole=0.4), use_container_width=True)
else:
    st.warning("조건에 맞는 데이터가 없습니다.")

# ----------------------------------------------------------------📄 메인 화면: 데이터 표 및 다운로드
st.markdown("---")
st.subheader("📋 상세 내역 리스트")
st.dataframe(filtered_df, use_container_width=True)

# 다운로드 버튼
import io
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    filtered_df.to_excel(writer, index=False)
st.download_button(
    label="📥 현재 데이터 엑셀 다운로드", data=output.getvalue(),
    file_name="discipline_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
