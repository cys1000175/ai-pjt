import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(layout="wide")
st.title("HR Dashboard 2018-2026")

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
            
            # 헤더 행 찾기 (번호가 있는 행)
            start_idx = 0
            for idx, row in df.iterrows():
                row_str = [str(x) for x in row.values if pd.notna(x)]
                if any("번호" in s for s in row_str):
                    start_idx = idx
                    break
            
            # 데이터 재로드 및 컬럼명 공백 제거
            data_df = excel.parse(sheet, skiprows=start_idx)
            data_df.columns = [str(c).strip() for c in data_df.columns]
            
            # 필수 기재 항목인 성명이 비어있는 행 제외
            name_col = [c for c in data_df.columns if "성명" in c or "성 명" in c]
            if name_col:
                data_df = data_df.dropna(subset=[name_col[0]])
            else:
                continue
                
            # 데이터 표준화 매핑 (안전하게 위치 기반 등으로 유연 처리)
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
        
        # 텍스트 정제 정규화
        final_df["Dept_Clean"] = final_df["Dept"].astype(str).apply(lambda x: x.split("\n")[0].split("(")[0].strip())
        final_df["Type_Clean"] = final_df["Type"].astype(str).apply(lambda x: next((t for t in ["해고", "강등", "정직", "감봉", "견책", "경고", "훈계", "권고"] if t in x), "기타"))
        
        # 년도 처리
        final_df["Year"] = pd.to_numeric(final_df["Year"], errors="coerce").fillna(2026).astype(int)
        final_df = final_df[(final_df["Year"] >= 2018) & (final_df["Year"] <= 2026)]
        
        final_df = final_df.sort_values(by=["Year", "Date"]).reset_index(drop=True)
        return final_df
    except:
        return pd.DataFrame()

df = load_all_data()

if df.empty:
    st.error("No data loaded. Please check data.xlsx file.")
else:
    # 필터 레이아웃
    st.write("### Filters")
    c1, c2, c3 = st.columns(3)
    with c1:
        f_yr = st.multiselect("Select Year", options=sorted(df["Year"].unique()))
    with c2:
        f_dp = st.multiselect("Select Department", options=sorted(df["Dept_Clean"].unique()))
    with c3:
        f_nm = st.multiselect("Select Name", options=sorted(df["Name"].unique()))
        
    f_df = df.copy()
    if f_yr:
        f_df = f_df[f_df["Year"].isin(f_yr)]
    if f_dp:
        f_df = f_df[f_df["Dept_Clean"].isin(f_dp)]
    if f_nm:
        f_df = f_df[f_df["Name"].isin(f_nm)]
        
    # KPI 요약 지표
    st.write("---")
    k1, k2, k3 = st.columns(3)
    k1.metric("Total Records", f"{len(f_df)} items")
    k2.metric("Unique Departments", f"{f_df['Dept_Clean'].nunique()} locations")
    k3.metric("Most Common Type", f"{f_df['Type_Clean'].value_counts().idxmax() if not f_df.empty else 'N/A'}")
    
    # 시각화 그래프
    st.write("---")
    st.write("### Statistical Charts")
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.write("Records by Division")
        st.plotly_chart(px.bar(f_df, x="Division", color="Division"), use_container_width=True)
        
        st.write("Records by Department Top 10")
        top_depts = f_df["Dept_Clean"].value_counts().head(10).reset_index()
        st.plotly_chart(px.bar(top_depts, x="Dept_Clean", y="count"), use_container_width=True)
        
    with col_right:
        st.write("Yearly Trend")
        trend = f_df.groupby("Year").size().reset_index(name="count")
        st.plotly_chart(px.line(trend, x="Year", y="count", markers=True), use_container_width=True)
        
        st.write("Discipline Type Distribution")
        st.plotly_chart(px.pie(f_df, names="Type_Clean", hole=0.3), use_container_width=True)
        
    # 데이터 테이블 및 다운로드
    st.write("---")
    st.write("### Integrated Data Master Table")
    show_cols = ["Year", "Division", "Date", "Dept", "Position", "Name", "Reason", "Type"]
    st.dataframe(f_df[show_cols], use_container_width=True, hide_index=True)
    
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        f_df[show_cols].to_excel(writer, index=False)
    st.download_button(label="Download Excel Report", data=output.getvalue(), file_name="hr_report.xlsx")
