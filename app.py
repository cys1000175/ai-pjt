import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
import os
import re

# 1. 페이지 기본 설정 및 디자인 레이아웃 정의
st.set_page_config(page_title="통합 징계 내역 관리 시스템", layout="wide")
st.title("📊 법인별·사업장별 징계 내역 통합 대시보드 (2018~2026)")
st.markdown("AI 공모전 제출용 다중 시트 자동 통합 및 데이터 정제 시스템 프로토타입입니다.")

# 2. 엑셀 데이터 로드 및 자동 정제(텍스트 표준화) 로직
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
            st.error(f"❌ 다중 시트 통합 중 에러 발생: {e}")
            return get_sample_data()
    else:
        st.info("💡 데이터 연동 대기 중입니다. 기본 샘플 데이터를 활성화합니다.")
        return get_sample_data()

def get_sample_data():
    return pd.DataFrame([
        {
            "번호": 1, "년도": 2026, "구분": "샘플법인", "소속": "서울본사 (미화)", "소속_정제": "서울본사",
            "직책": "대리", "성명": "홍길동", "징계 사유": "엑
