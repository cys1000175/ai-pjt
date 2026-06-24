import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(layout="wide")
st.title("인사 징계 내역 통합 관리 대시보드")

# [보안 설정] 원하는 비밀번호를 여기에 지정하세요.
ADMIN_PASSWORD = "1234"

# 비밀번호 확인을 위한 세션 상태 초기화
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# 비밀번호 입력 창 UI (인증되지 않은 경우에만 표시)
if not st.session_state.authenticated:
    st.write("### 🔒 보안 잠금 시스템")
    st.info("본 대시보드는 민감한 개인정보를 포함하고 있습니다. 비밀번호를 입력해 주세요.")
    
    input_pw = st.text_input("비밀번호 입력", type="password")
    if st.button("인증하기"):
        if input_pw == ADMIN_PASSWORD:
            st.session_state.authenticated = True
            st.success("🔒 인증 성공! 대시보드를 로드합니다.")
            st.rerun()
        else:
            st.error("❌ 비밀번호가 일치하지 않습니다. 다시 시도해 주세요.")
    st.stop() # 인증 전에는 하위 코드가 절대 실행되지 않도록 강제 차단

FILE_NAME = "data.xlsx"

@st.cache_data
def load_all_data():
    if not os.path.exists(FILE_NAME):
        return pd.DataFrame()
    try:
        excel = pd.ExcelFile(FILE_NAME)
        all_sheets = []
        for sheet in excel.sheet_names:
            df = excel.parse(sheet, header=None)
            if df.empty:
                continue
            
            start_idx = 0
            for idx, row in df.iterrows():
                row_str = [str(x) for x in row.values if pd.notna(x)]
                if any("번호" in s for s in row_str):
                    start_idx = idx
                    break
            
            data_df = excel.parse(sheet, skiprows=start_idx)
            data_df.columns = [str(c).strip() for c in data_df.columns]
            
            name_col = [c for c in data_df.columns if "성명" in c or "성 명" in c]
            if name_col:
                data_df = data_df.dropna(subset=[name_col[0]])
            else:
                continue
                
            res = pd.DataFrame()
            
            def get_val(keywords, default=""):
                for c in data_df.columns:
                    if any(k in c for k in keywords):
                        return data_df[c].fillna(default)
                return default

            res["Year"] = get_val(["년도", "연도"], sheet.replace("년", "").replace("원", "").strip())
            res["Division"] = get_val(["구분"])
            res["Date"] = get_val(["일 자", "일자", "징계일"])
            res["Dept"] = get_val(["소속"])
            res["Position"] = get_val(["직책", "직급"])
            res["Name"] = get_val(["성명", "성 명"])
            res["Reason"] = get_val(["사유", "징계 사유", "징계사유"])
            res["Type"] = get_val(["종류", "징계 종류", "징계종류"])
            
            all_sheets.append(res)
            
        if not all_sheets:
            return pd.DataFrame()
            
        final_df = pd.concat(all_sheets, ignore_index=True)
        final_df = final_df.dropna(subset=["Name"])
        final_df["Name"] = final_df["Name"].astype(str).str.strip()
        final_df = final_df[final_df["Name"] != ""]
        
        final_df["Dept_Clean"] = final_df["Dept"].astype(str).apply(lambda x: x.split("\n")[0].split("(")[0].strip())
        final_df["Type_Clean"] = final_df["Type"].astype(str).apply(lambda x: next((t for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고"] if t in x), "기타"))
        
        final_df["Year"] = pd.to_numeric(final_df["Year"], errors="coerce").fillna(2026).astype(int)
        final_df = final_df[(final_df["Year"] >= 2018) & (final_df["Year"] <= 2026)]
        
        final_df = final_df.sort_values(by=["Year", "Date"]).reset_index(drop=True)
        return final_df
    except:
        return pd.DataFrame()

df = load_all_data()

if df.empty:
    st.error("데이터를 로드하지 못했습니다. data.xlsx 파일을 확인해 주세요.")
else:
    st.write("### 🔍 데이터 통합 검색 필터")
    c1, c2, c3 = st.columns(3)
    with c1:
        f_yr = st.multiselect("연도 선택 (비워두면 전체 조회)", options=sorted(df["Year"].unique()))
    with c2:
        f_dp = st.multiselect("소속 사업장 선택 (비워두면 전체 조회)", options=sorted(df["Dept_Clean"].unique()))
    with c3:
        f_nm = st.multiselect("성명 검색 (비워두면 전체 조회)", options=sorted(df["Name"].unique()))
        
    f_df = df.copy()
    if f_yr:
        f_df = f_df[f_df["Year"].isin(f_yr)]
    if f_dp:
        f_df = f_df[f_df["Dept_Clean"].isin(f_dp)]
    if f_name: # 💥 변수명 정합성 유지
        f_df = f_df[f_df["Name"].isin(f_nm)]
        
    st.write("---")
    k1, k2, k3 = st.columns(3)
    k1.metric("총 누적 발생 건수", f"{len(f_df)} 건")
    k2.metric("분석 대상 사업장 수", f"{f_df['Dept_Clean'].nunique()} 개소")
    
    most_type = f_df['Type_Clean'].value_counts().idxmax() if not f_df.empty else '없음'
    k3.metric("가장 빈번한 징계 종류", f"{most_type}")
    
    st.write("---")
    st.write("### 📊 실시간 데이터 통계 시각화")
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.write("#### 🏢 구분별 발생 건수")
        st.plotly_chart(px.bar(f_df, x="Division", color="Division", labels={"Division": "위원회 구분", "count": "건수"}), use_container_width=True)
        
        st.write("#### 📍 소속 사업장별 발생 TOP 10")
        top_depts = f_df["Dept_Clean"].value_counts().head(10).reset_index()
        top_depts.columns = ["사업장명", "건수"]
        st.plotly_chart(px.bar(top_depts, x="사업장명", y="건수", color="사업장명"), use_container_width=True)
        
    with col_right:
        st.write("#### 📅 연도별 장기 추이 트렌드 (2018-2026)")
        trend = f_df.groupby("Year").size().reset_index(name="건수")
        st.plotly_chart(px.line(trend, x="Year", y="건수", markers=True, labels={"Year": "연도"}), use_container_width=True)
        
        st.write("#### ⚖️ 전체 징계 유형별 비율 현황")
        st.plotly_chart(px.pie(f_df, names="Type_Clean", hole=0.3), use_container_width=True)
        
    st.write("---")
    st.write("### 📋 통합 상세 내역 마스터 테이블")
    
    show_df = f_df[["Year", "Division", "Date", "Dept", "Position", "Name", "Reason", "Type"]].copy()
    show_df.columns = ["년도", "구분", "일자", "소속", "직책", "성명", "징계 사유", "징계종류"]
    
    st.dataframe(show_df, use_container_width=True, hide_index=True)
    
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        show_df.to_excel(writer, index=False)
    st.download_button(label="📥 통합 정제 데이터 다운로드 (Excel)", data=output.getvalue(), file_name="HR_integrated_report.xlsx")
