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
    val = re.sub(r'\s*\(.*?\)\s*', '', val)
    return val.strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
            combined_rows = []
            
            for sheet_name, raw_df in all_sheets.items():
                header_row_idx = None
                
                for idx, row in raw_df.iterrows():
                    row_str = [str(x) for x in row.values]
                    if any("번호" in s or "년도" in s or "구분" in s for s in row_str):
                        header_row_idx = idx
                        break
                
                if header_row_idx is not None:
                    sheet_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, skiprows=header_row_idx)
                else:
                    sheet_df = raw_df.copy()
                
                sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                
                if sheet_df.empty:
                    continue
                    
                for col in REQUIRED_COLUMNS:
                    if col not in sheet_df.columns:
                        matched_col = [c for c in sheet_df.columns if col in c or c in col]
                        if matched_col:
                            sheet_df[col] = sheet_df[matched_col[0]]
                        else:
                            sheet_df[col] = ""
                
                sheet_df = sheet_df[REQUIRED_COLUMNS]
                combined_rows.append(sheet_df)
            
            if combined_rows:
                total_df = pd.concat(combined_rows, ignore_index=True)
            else:
                return get_sample_data()
                
            total_df = total_df.dropna(subset=["성명", "징계종류"], how="all")
            total_df["구분"] = total_df["구분"].fillna("미지정").astype(str).str.strip()
            total_df["구분"] = total_df["구분"].apply(lambda x: "미지정" if x == "" or x == "nan" else x)
            
            total_df["소속_정제"] = total_df["소속"].apply(clean_location_name)
            total_df["징계종류_정제"] = total_df["징계종류"].apply(clean_discipline_type)
            
            total_df["년도"] = pd.to_numeric(total_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            total_df["성명"] = total_df["성명"].fillna("미상").astype(str).str.strip()
            
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

# ----------------------------------------------------------------🔍 메인 화면: 동적 필터 (안전 잠금 강화)
st.subheader("🔍 데이터 통합 검색 필터")

if not df.empty:
    col1, col2, col3 = st.columns(3)
    with col1:
        # 아무것도 선택 안했을 때는 전체 조회가 되도록 default 세팅을 최적화했습니다.
        f_year = st.multiselect("년도 선택 (비워두면 전체 조회)", options=sorted(list(df["년도"].unique())))
    with col2:
        f_dept = st.multiselect("소속 사업장 선택 (비워두면 전체 조회)", options=sorted(list(df["소속_정제"].unique())))
    with col3:
        f_name = st.multiselect("성명 검색 (비워두면 전체 조회)", options=sorted(list(df["성명"].unique())))

    # [🔥 해결 포인트] 필터가 비어 있으면(선택을 안 하면) 해당 조건은 패스(전체선택)하도록 로직 고도화
    filtered_df = df.copy()
    if f_year:
        filtered_df = filtered_df[filtered_df["년도"].isin(f_year)]
    if f_dept:
        filtered_df = filtered_df[filtered_df["소속_정제"].isin(f_dept)]
    if f_name:
        filtered_df = filtered_df[filtered_df["성명"].isin(f_name)]
else:
    filtered_df = df

# ----------------------------------------------------------------📊 메인 화면: 시각화 차트
st.markdown("---")
st.subheader("📊 연도별·사업장별 통합 실시간 트렌드")

if not filtered_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🏢 구분(법인)별 누적 발생 건수")
        st.plotly_chart(px.bar(filtered_df, x="구분", color="구분", labels={"구분": "법인 구분", "count": "건수"}), use_container_width=True)
        
        st.markdown("#### 📍 소속(사업장)별 징계 지표")
        st.plotly_chart(px.histogram(filtered_df, x="소속_정제", color="징계종류_정제", barmode="stack", labels={"소속_정제": "소속 사업장", "징계종류_정제": "징계 종류"}), use_container_width=True)
    with c2:
        st.markdown("#### 📅 2018-2026 연도별 발생 추이 추적")
        year_trend = filtered_df.groupby("년도").size().reset_index(name="건수").sort_values("년도")
        st.plotly_chart(px.line(year_trend, x="년도", y="건수", text="건수", markers=True), use_container_width=True)
        
        st.markdown("#### ⚖️ 전체 징계종류 비율 분석")
        st.plotly_chart(px.pie(filtered_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("⚠️ 선택하신 필터 조합(년도+소속+성명)에 동시에 만족하는 데이터가 데이터베이스에 없습니다. 필터 선택을 가볍게 클릭하여 해제해 보세요!")

# ----------------------------------------------------------------📄 메인 화면: 데이터 표 및 다운로드
st.markdown("---")
st.subheader("📋 2018~2026 전 시트 통합 상세 내역 리스트")

display_cols = [c for c in REQUIRED_COLUMNS if c in filtered_df.columns]
if filtered_df.empty:
    st.info("조건을 만족하는 징계 내역 데이터가 없습니다. 상단 필터를 비우면 전체 내역이 다시 나타납니다.")
else:
    display_df = filtered_df[display_cols]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        display_df.to_excel(writer, index=False)
    st.
