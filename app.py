import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os
import re

st.set_page_config(page_title="통합 징계 관리", layout="wide")
st.title("📊 법인별·사업장별 징계 내역 통합 대시보드 (2018~2026)")
st.markdown("AI 공모전 제출용 다중 시트 자동 통합 및 데이터 정제 시스템입니다.")

EXCEL_FILE = "data.xlsx"
REQUIRED_COLUMNS = ["번호", "년도", "구분", "소속", "직책", "성명", "징계 사유", "징계종류", "징계일"]

def clean_discipline_type(val):
    if pd.isna(val) or not isinstance(val, str): return "기타"
    val = val.strip()
    for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고사직"]:
        if t in val: return t
    return "기타"

def clean_location_name(val):
    if pd.isna(val) or not isinstance(val, str): return "기타 사업장"
    return re.sub(r'\s*\(.*?\)\s*', '', val.replace('\n', ' ')).strip()

def load_initial_data():
    if os.path.exists(EXCEL_FILE):
        try:
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
            combined_rows = []
            for sheet_name, raw_df in all_sheets.items():
                hdr_idx = 0
                for idx, row in raw_df.iterrows():
                    if any(k in [str(x) for x in row.values] for k in ["번호", "년도", "구분"]):
                        hdr_idx = idx
                        break
                sheet_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, skiprows=hdr_idx)
                sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                if sheet_df.empty: continue
                for col in REQUIRED_COLUMNS:
                    if col not in sheet_df.columns:
                        m = [c for c in sheet_df.columns if col in c or c in col]
                        sheet_df[col] = sheet_df[m[0]] if m else ""
                combined_rows.append(sheet_df[REQUIRED_COLUMNS])
            
            if not combined_rows: return get_sample_data()
            total_df = pd.concat(combined_rows, ignore_index=True).dropna(subset=["성명", "징계종류"], how="all")
            total_df["구분"] = total_df["구분"].fillna("미지정").astype(str).str.strip().replace("", "미지정")
            total_df["소속_정제"] = total_df["소속"].apply(clean_location_name)
            total_df["징계종류_정제"] = total_df["징계종류"].apply(clean_discipline_type)
            total_df["년도"] = pd.to_numeric(total_df["년도"], errors='coerce').fillna(datetime.date.today().year).astype(int)
            total_df["성명"] = total_df["성명"].fillna("미상").astype(str).str.strip()
            total_df = total_df.sort_values(by=["년도", "징계일"]).reset_index(drop=True)
            total_df["번호"] = total_df.index + 1
            return total_df
        except Exception as e:
            st.error(f"❌ 에러 발생: {e}")
            return get_sample_data()
    return get_sample_data()

def get_sample_data():
    return pd.DataFrame([{c: "샘플" for c in REQUIRED_COLUMNS} | {"년도": 2026, "소속_정제": "샘플", "징계종류_정제": "기타"}])

if 'discipline_data' not in st.session_state:
    st.session_state.discipline_data = load_initial_data()

df = st.session_state.discipline_data

st.sidebar.header("➕ 신규 징계 내역 입력")
with st.sidebar.form(key="input_form", clear_on_submit=True):
    div_opts = list(df["구분"].dropna().unique()) if not df.empty else ["기본 법인"]
    input_division = st.selectbox("구분(법인) 선택", div_opts + ["기타"] if "기타" not in div_opts else div_opts)
    input_year = st.number_input("년도", min_value=2018, max_value=2030, value=datetime.date.today().year)
    input_date = st.date_input("징계일 선택", datetime.date.today())
    input_dept = st.text_input("소속(사업장명)")
    input_position = st.text_input("직책")
    input_name = st.text_input("성명")
    input_type = st.selectbox("징계종류", ["견책", "감봉", "정직", "강등", "해고", "경고", "훈계", "권고사직"])
    input_reason = st.text_area("징계 사유")
    submit_button = st.form_submit_button(label="대시보드에 추가")

if submit_button and input_dept and input_name and input_reason:
    new_row = {
        "번호": len(df) + 1, "년도": int(input_year), "구분": input_division, "소속": input_dept,
        "소속_정제": clean_location_name(input_dept), "직책": input_position, "성명": input_name,
        "징계 사유": input_reason, "징계종류": input_type, "징계종류_정제": clean_discipline_type(input_type),
        "징계일": input_date.strftime("%Y-%m-%d")
    }
    st.session_state.discipline_data = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    st.sidebar.success("✅ 반영되었습니다!")
    st.rerun()

df = st.session_state.discipline_data

st.subheader("🔍 데이터 통합 검색 필터")
if not df.empty:
    col1, col2, col3 = st.columns(3)
    with col1: f_year = st.multiselect("년도 선택 (비워두면 전체)", options=sorted(list(df["년도"].unique())))
    with col2: f_dept = st.multiselect("소속 사업장 선택 (비워두면 전체)", options=sorted(list(df["소속_정제"].unique())))
    with col3: f_name = st.multiselect("성명 검색 (비워두면 전체)", options=sorted(list(df["성명"].unique())))
    
    filtered_df = df.copy()
    if f_year: filtered_df = filtered_df[filtered_df["년도"].isin(f_year)]
    if f_dept: filtered_df = filtered_df[filtered_df["소속_정제"].isin(f_dept)]
    if f_name: filtered_df = filtered_df[filtered_df["성명"].isin(f_name)]
else:
    filtered_df = df

st.markdown("---")
st.subheader("📊 연도별·사업장별 통합 실시간 트렌드")
if not filtered_df.empty:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🏢 구분(법인)별 누적 발생 건수")
        st.plotly_chart(px.bar(filtered_df, x="구분", color="구분"), use_container_width=True)
        st.markdown("#### 📍 소속(사업장)별 징계 지표")
        st.plotly_chart(px.histogram(filtered_df, x="소속_정제", color="징계종류_정제", barmode="stack"), use_container_width=True)
    with c2:
        st.markdown("#### 📅 연도별 발생 추이 추적 (2018-2026)")
        yt = filtered_df.groupby("년도").size().reset_index(name="건수")
        st.plotly_chart(px.line(yt, x="년도", y="건수", text="건수", markers=True), use_container_width=True)
        st.markdown("#### ⚖️ 전체 징계종류 비율 분석")
        st.plotly_chart(px.pie(filtered_df, names="징계종류_정제", hole=0.4), use_container_width=True)
else:
    st.warning("⚠️ 선택하신 조건에 만족하는 내역이 없습니다.")

st.markdown("---")
st.subheader("📋 통합 상세 내역 리스트")
if not filtered_df.empty:
    display_df = filtered_df[[c for c in REQUIRED_COLUMNS if c in filtered_df.columns]]
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: display_df.to_excel(w, index=False)
    st.download_button(label="📥 통합 본 데이터 엑셀 다운로드", data=output.getvalue(), file_name="integrated_report.xlsx")
else:
    st.info("데이터가 없습니다.")
