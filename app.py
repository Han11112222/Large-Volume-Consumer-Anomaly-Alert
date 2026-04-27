import io
import json
import os
import re
import random
from pathlib import Path
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib as mpl
import plotly.graph_objects as go
import pydeck as pdk
import requests
import streamlit as st
from github import Github


# ─────────────────────────────────────────────────────────
# 기본 설정
# ─────────────────────────────────────────────────────────
def set_korean_font():
    ttf = Path(__file__).parent / "NanumGothic-Regular.ttf"
    if ttf.exists():
        try:
            mpl.font_manager.fontManager.addfont(str(ttf))
            mpl.rcParams["font.family"] = "NanumGothic"
            mpl.rcParams["axes.unicode_minus"] = False
        except Exception:
            pass

set_korean_font()
st.set_page_config(page_title="대용량 수요처 이상 감지 대시보드", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"

# ─────────────────────────────────────────────────────────
# 코멘트 DB 저장 (비밀번호 없음)
# ─────────────────────────────────────────────────────────
COMMENT_DB_FILE = "report_comments_db.json"
REPO_NAME = "Han11112222/quarterly-sales-report"

def load_comments_db():
    if os.path.exists(COMMENT_DB_FILE):
        try:
            with open(COMMENT_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_comments_db(db_data):
    with open(COMMENT_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db_data, f, ensure_ascii=False, indent=4)
    try:
        if "GITHUB_TOKEN" in st.secrets:
            token = st.secrets["GITHUB_TOKEN"]
            g = Github(token)
            repo = g.get_repo(REPO_NAME)
            content_string = json.dumps(db_data, ensure_ascii=False, indent=4)
            try:
                contents = repo.get_contents(COMMENT_DB_FILE)
                repo.update_file(contents.path, "Update comments via Streamlit App", content_string, contents.sha)
            except:
                repo.create_file(COMMENT_DB_FILE, "Create comments db via Streamlit App", content_string)
    except Exception:
        pass

def render_comment_section(title, db_key, curr_db, comments_db, height, placeholder, widget_key):
    st.markdown(f"**{title}**")
    saved_text = curr_db.get(db_key, None)
    
    if saved_text is not None:
        url_pattern = re.compile(r'(https?://[^\s]+)')
        linked_text = url_pattern.sub(r'<a href="\1" target="_blank" style="color: #2563eb; text-decoration: underline; font-weight: bold;">\1</a>', saved_text)
        formatted_text = linked_text.replace('\n', '<br>')
        st.markdown(
            f"""
            <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-left: 4px solid #1f77b4; padding: 15px; border-radius: 4px; color: #1e40af; font-size: 14.5px; line-height: 1.6; margin-bottom: 10px;">
                {formatted_text}
            </div>
            """, unsafe_allow_html=True
        )
        
        with st.expander("📝 코멘트 수정/삭제"):
            new_text = st.text_area("내용 수정", value=saved_text, height=height, key=f"edit_ta_{widget_key}", label_visibility="collapsed")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 수정 내용 저장", key=f"edit_save_{widget_key}", use_container_width=True):
                    curr_db[db_key] = new_text
                    save_comments_db(comments_db)
                    st.rerun()
            with col2:
                if st.button("🗑️ 코멘트 삭제", key=f"del_{widget_key}", use_container_width=True):
                    curr_db.pop(db_key, None)
                    save_comments_db(comments_db)
                    st.rerun()
    else:
        input_text = st.text_area("내용 입력", height=height, placeholder=placeholder, key=f"ta_{widget_key}", label_visibility="collapsed")
        if st.button("💾 이 코멘트 저장", key=f"save_{widget_key}"):
            curr_db[db_key] = input_text
            save_comments_db(comments_db)
            st.rerun()


# ─────────────────────────────────────────────────────────
# 데이터 전처리 유틸
# ─────────────────────────────────────────────────────────
USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "일반용": "영업용",
    "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용",
    "수송용(CNG)": "수송용", "수송용(BIO)": "수송용",
    "열병합용": "열병합", "열병합용1": "열병합", "열병합용2": "열병합",
    "연료전지용": "연료전지", "열전용설비용": "열전용설비용",
}

COLOR_ACT = "rgba(0, 150, 255, 1)"
COLOR_PREV = "rgba(190, 190, 190, 1)"
COLOR_ALARM = [211, 47, 47, 200]

def clean_korean_finance_number(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(",", "").strip()
    if not s: return 0.0
    if s.endswith("-"): s = "-" + s[:-1]
    elif s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    s = re.sub(r"[^\d\.-]", "", s)
    try: return float(s)
    except: return 0.0

def fmt_num_safe(v) -> str:
    if pd.isna(v): return "-"
    try: return f"{float(v):,.0f}"
    except Exception: return "-"

def center_style(styler):
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([
        dict(selector="th", props=[("text-align", "center"), ("vertical-align", "middle"), ("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")]),
        dict(selector="thead th", props=[("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")]),
        dict(selector="tbody tr th", props=[("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")])
    ])
    return styler

def highlight_subtotal(s):
    is_subtotal = s.astype(str).str.contains('💡 소계|💡 총계|💡 합계')
    return ['background-color: #1e3a8a; color: #ffffff; font-weight: bold;' if is_subtotal.any() else '' for _ in s]

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns: out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def keyword_group(col: str) -> Optional[str]:
    c = str(col)
    if "열병합" in c: return "열병합"
    if "연료전지" in c: return "연료전지"
    if "수송용" in c: return "수송용"
    if "열전용" in c: return "열전용설비용"
    if c in ["산업용"]: return "산업용"
    if c in ["일반용"]: return "영업용"
    if any(k in c for k in ["취사용", "난방용", "자가열"]): return "가정용"
    if any(k in c for k in ["업무", "냉방", "주한미군"]): return "업무용"
    return None

def make_long(plan_df: pd.DataFrame, actual_df: pd.DataFrame) -> pd.DataFrame:
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)
    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue
            group = USE_COL_TO_GROUP.get(col)
            if group is None: group = keyword_group(col)
            if group is None: continue
            base = df[["연", "월"]].copy()
            base["그룹"] = group
            base["용도"] = col
            base["계획/실적"] = label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)
    if not records: return pd.DataFrame(columns=["연", "월", "그룹", "용도", "계획/실적", "값"])
    long_df = pd.concat(records, ignore_index=True)
    long_df = long_df.dropna(subset=["연", "월"])
    long_df["연"] = long_df["연"].astype(int)
    long_df["월"] = long_df["월"].astype(int)
    return long_df

def load_all_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    needed = ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]
    out: Dict[str, pd.DataFrame] = {}
    for name in needed:
        if name in xls.sheet_names: out[name] = xls.parse(name)
    return out

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    long_dict: Dict[str, pd.DataFrame] = {}
    if ("계획_부피" in sheets) and ("실적_부피" in sheets):
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if ("계획_열량" in sheets) and ("실적_열량" in sheets):
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

def load_safe_csv(file_bytes) -> pd.DataFrame:
    encodings = ["utf-8-sig", "cp949", "utf-8", "euc-kr"]
    for enc in encodings:
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, thousands=',')
            df.columns = df.columns.str.strip() 
            return df
        except Exception:
            pass
    return pd.DataFrame()

def get_coord_from_df(address: str, coord_df: pd.DataFrame) -> Tuple[float, float]:
    if pd.isna(address) or not str(address).strip():
        return None, None
        
    if not coord_df.empty and len(coord_df.columns) >= 3:
        clean_addr = re.sub(r'\(.*?\)', '', str(address))
        clean_addr = clean_addr.split(',')[0].strip()
        clean_addr_no_space = clean_addr.replace(" ", "")
        
        if clean_addr_no_space:
            addr_col = coord_df.columns[0]
            lat_col = coord_df.columns[1]
            lon_col = coord_df.columns[2]
            
            coord_addrs = coord_df[addr_col].astype(str).str.replace(" ", "", regex=False)
            mask = coord_addrs.str.contains(re.escape(clean_addr_no_space), na=False)
            match = coord_df[mask]
            
            if not match.empty:
                try:
                    return float(match.iloc[0][lat_col]), float(match.iloc[0][lon_col])
                except:
                    pass
                    
    lat = 35.8714 + random.uniform(-0.06, 0.06)
    lon = 128.6014 + random.uniform(-0.06, 0.06)
    return lat, lon


# ─────────────────────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────────────────────
st.title("📊 대용량 수요처 이상 감지 대시보드")

with st.sidebar:
    st.header("📂 데이터 및 설정")

    st.subheader("1. 판매량 데이터 (요약/엑셀)")
    src_sales = st.radio("판매량 데이터 소스", ["레포 파일 사용", "엑셀 업로드(.xlsx)"], index=0, key="rpt_sales_src")
    excel_bytes = None
    if src_sales == "엑셀 업로드(.xlsx)":
        up_sales = st.file_uploader("판매량(계획_실적).xlsx 형식", type=["xlsx"], key="rpt_sales_uploader")
        if up_sales is not None: excel_bytes = up_sales.getvalue()
    else:
        path_sales = Path(__file__).parent / DEFAULT_SALES_XLSX
        if path_sales.exists(): excel_bytes = path_sales.read_bytes()
    
    st.markdown("---")

    st.subheader("2. 업종별 데이터 (상세/CSV)")
    src_csv = st.radio("업종별 데이터 소스", ["레포 파일 사용", "CSV 업로드(.csv)"], index=0, key="csv_src")
    if src_csv == "CSV 업로드(.csv)":
        up_csvs = st.file_uploader("가정용외_*.csv 형식 (다중 업로드 가능)", type=["csv"], accept_multiple_files=True, key="csv_uploader")
        if up_csvs:
            df_list = []
            for f in up_csvs:
                df = load_safe_csv(f.getvalue())
                if not df.empty:
                    df_list.append(df)
            if df_list: st.session_state['merged_csv_df'] = pd.concat(df_list, ignore_index=True)
        else:
            if 'merged_csv_df' in st.session_state: del st.session_state['merged_csv_df']

    st.markdown("---")
    st.subheader("🗺️ 3. 지도 위경도 데이터 (CSV)")
    src_coord = st.radio("위경도 데이터 소스", ["레포 파일(깃허브) 사용", "CSV 업로드(.csv)"], index=0, key="coord_src")
    
    coord_df = pd.DataFrame()
    if src_coord == "CSV 업로드(.csv)":
        up_coord = st.file_uploader("위경도 매핑 파일 업로드 (address_with_latlon.csv)", type=["csv"], key="coord_uploader")
        if up_coord:
            coord_df = load_safe_csv(up_coord.getvalue())
    else:
        coord_path = Path(__file__).parent / "address_with_latlon.csv"
        if coord_path.exists():
            coord_df = load_safe_csv(coord_path.read_bytes())
        else:
            github_csv_url = "https://raw.githubusercontent.com/Han11112222/quarterly-sales-report/main/address_with_latlon.csv"
            try:
                res = requests.get(github_csv_url, timeout=5)
                if res.status_code == 200:
                    coord_df = load_safe_csv(res.content)
            except:
                pass


# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────

long_dict_rpt: Dict[str, pd.DataFrame] = {}
if excel_bytes is not None:
    sheets_rpt = load_all_sheets(excel_bytes)
    long_dict_rpt = build_long_dict(sheets_rpt)
    
df_csv = pd.DataFrame()

if src_csv == "레포 파일 사용":
    repo_dir = Path(__file__).parent
    all_csvs = list(repo_dir.glob("*가정용외*.csv")) + list(repo_dir.glob("가정용외*.csv"))
    all_csvs = list(set(all_csvs)) 
    csv_list = []
    for p in all_csvs:
        try:
            temp_df = pd.read_csv(p, encoding="utf-8-sig", thousands=',')
            temp_df.columns = temp_df.columns.str.strip()
            csv_list.append(temp_df)
        except:
            try: 
                temp_df = pd.read_csv(p, encoding="cp949", thousands=',')
                temp_df.columns = temp_df.columns.str.strip()
                csv_list.append(temp_df)
            except: pass
    if csv_list: df_csv = pd.concat(csv_list, ignore_index=True)

if df_csv.empty and 'merged_csv_df' in st.session_state:
    df_csv = st.session_state['merged_csv_df'].copy()
    
if not df_csv.empty:
    if "사용량(mj)" in df_csv.columns: df_csv["사용량(mj)"] = df_csv["사용량(mj)"].apply(clean_korean_finance_number)
    if "사용량(m3)" in df_csv.columns: df_csv["사용량(m3)"].apply(clean_korean_finance_number)
        
comments_db = load_comments_db()
        
rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

for idx, rpt_tab in enumerate(rpt_tabs):
    with rpt_tab:
        if idx == 0:
            df_long_rpt = long_dict_rpt.get("열량", pd.DataFrame())
            unit_str = "GJ"
            val_col = "사용량(mj)"
            key_sfx = "_gj"
        else:
            df_long_rpt = long_dict_rpt.get("부피", pd.DataFrame())
            unit_str = "천m³"
            val_col = "사용량(m3)"
            key_sfx = "_vol"

        st.markdown(f"#### 📅 기준 일자 설정") 
        
        years_available = [2024, 2025, 2026]
        default_y_index = len(years_available) - 1
        default_m_index = 2 
        
        if not df_long_rpt.empty:
            years_available = sorted(df_long_rpt["연"].unique().tolist())
            actual_data = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["값"] > 0)]
            if not actual_data.empty:
                max_year = actual_data["연"].max()
                max_month = actual_data[actual_data["연"] == max_year]["월"].max()
                default_y_index = years_available.index(max_year) if max_year in years_available else len(years_available) - 1
                default_m_index = int(max_month - 1)
                
        df_csv_tab = df_csv.copy()
        if not df_csv_tab.empty:
            df_csv_tab["연_csv"] = 2026
            df_csv_tab["월_csv"] = 3
            
            if unit_str == "GJ" and "사용량(mj)" in df_csv_tab.columns:
                df_csv_tab["사용량(mj)"] = pd.to_numeric(df_csv_tab["사용량(mj)"], errors="coerce").fillna(0) / 1000.0
            elif unit_str == "천m³" and "사용량(m3)" in df_csv_tab.columns:
                df_csv_tab["사용량(m3)"] = pd.to_numeric(df_csv_tab["사용량(m3)"], errors="coerce").fillna(0) / 1000.0
                
            df_csv_tab["날짜_파싱"] = pd.to_datetime("2026-03-01")
            date_col = None
            for c in ["청구년월", "매출년월", "년월", "기준년월"]:
                if c in df_csv_tab.columns:
                    date_col = c
                    break
                    
            if date_col:
                try:
                    parsed = pd.to_datetime(df_csv_tab[date_col], format="%b-%y", errors="coerce")
                    mask = parsed.isna()
                    if mask.any():
                        parsed.loc[mask] = pd.to_datetime(df_csv_tab.loc[mask, date_col], format="%Y%m", errors="coerce")
                    mask = parsed.isna()
                    if mask.any():
                        parsed.loc[mask] = pd.to_datetime(df_csv_tab.loc[mask, date_col], errors="coerce")
                    
                    df_csv_tab["날짜_파싱"] = parsed.fillna(pd.to_datetime("2026-03-01"))
                except Exception:
                    pass

            df_csv_tab["연_csv"] = df_csv_tab["날짜_파싱"].dt.year
            df_csv_tab["월_csv"] = df_csv_tab["날짜_파싱"].dt.month
        
        c_y, c_m, c_agg, c_empty = st.columns([1, 1, 2, 1])
        with c_y:
            sel_year_rpt = st.selectbox("기준 연도", years_available, index=default_y_index, key=f"rpt_yr{key_sfx}")
        with c_m:
            sel_month_str = st.selectbox("기준 월", [f"{m}월" for m in range(1, 13)], index=default_m_index, key=f"rpt_mo{key_sfx}")
        with c_agg:
            agg_mode = st.radio("집계 기준", ["당월 실적", "누적 실적 (1월~당월)"], index=0, horizontal=True, key=f"agg_mode_{key_sfx}")
        
        max_month = int(sel_month_str.replace("월", "")) 
        report_db_key = f"{sel_year_rpt}_{max_month}M_{unit_str}_yoy_only"
        
        if report_db_key not in comments_db: comments_db[report_db_key] = {}
        curr_db = comments_db[report_db_key]
        
        st.markdown("<hr style='margin: 10px 0 30px 0;'>", unsafe_allow_html=True)

        # ─────────────────────────────────────────────────────────
        # 1. 이상 감지 업체 지도 모니터링
        # ─────────────────────────────────────────────────────────
        st.markdown(f"### 🗺️ 1. 대용량 수요처 이상 감지 모니터링 지도 <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
        st.caption("※ YoY 기준 5% 이상 사용량이 하락한 업체를 지도에 마커로 표시하여 현장 방문을 유도합니다.")
        
        st.markdown("""
        <div style='background-color: #f1f3f5; padding: 12px; border-radius: 6px; margin-bottom: 15px; font-size: 14px;'>
            <b>💡 지도 마커(알람) 3단계 구분 안내</b><br>
            • <b>심각 (20% 이상 하락)</b> : 가장 크고 진한 색상의 마커<br>
            • <b>경계 (10% 이상 하락)</b> : 중간 크기와 중간 농도의 마커<br>
            • <b>주의 (5% 이상 하락)</b> : 작고 연한 색상의 마커<br>
            <span style='font-size: 12px; color: #555;'>※ 산업용은 붉은색(🔴), 업무용은 푸른색(🔵) 계열로 표시됩니다.</span>
        </div>
        """, unsafe_allow_html=True)
        
        map_c1, map_c2 = st.columns(2)
        with map_c1:
            map_usage = st.radio("📍 지도에 표시할 용도 선택", ["산업용", "업무용"], index=0, horizontal=True, key=f"map_radio_{key_sfx}")
        with map_c2:
            map_style_ui = st.radio("📍 지도 배경 테마", ["다크 모드 (기본)", "일반 도로 지도"], index=0, horizontal=True, key=f"map_style_{key_sfx}")
        
        deck_map_style = "dark" if map_style_ui == "다크 모드 (기본)" else "road"
        
        if not df_csv_tab.empty and "도로명주소" in df_csv_tab.columns and "고객명" in df_csv_tab.columns and val_col in df_csv_tab.columns and "용도" in df_csv_tab.columns:
            if agg_mode == "누적 실적 (1월~당월)":
                df_map_base = df_csv_tab[df_csv_tab["월_csv"] <= max_month].copy()
            else:
                df_map_base = df_csv_tab[df_csv_tab["월_csv"] == max_month].copy()
            
            if not df_map_base.empty:
                if map_usage == "산업용":
                    df_map_filtered = df_map_base[df_map_base["용도"] == "산업용"].copy()
                else: 
                    if "상품명" in df_map_base.columns:
                        prod_s = df_map_base["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                        mask = (df_map_base["용도"] == "업무용") | (prod_s.isin(["냉난방용(업무)", "업무난방용", "주한미군"]))
                        df_map_filtered = df_map_base[mask].copy()
                    else:
                        df_map_filtered = df_map_base[df_map_base["용도"] == "업무용"].copy()
                
                df_map_filtered["용도_태그"] = f"[{map_usage}]"

                map_curr = df_map_filtered[df_map_filtered["연_csv"] == sel_year_rpt].groupby(["고객명", "도로명주소", "용도_태그"], as_index=False)[val_col].sum().rename(columns={val_col: "당해년도"})
                map_prev = df_map_filtered[df_map_filtered["연_csv"] == sel_year_rpt - 1].groupby(["고객명", "도로명주소", "용도_태그"], as_index=False)[val_col].sum().rename(columns={val_col: "전년도"})
                
                if not map_curr.empty and not map_prev.empty:
                    df_map_merged = pd.merge(map_curr, map_prev, on=["고객명", "도로명주소", "용도_태그"], how="inner").fillna(0)
                    
                    df_map_merged["증감률(%)"] = np.where(df_map_merged["전년도"] > 0, ((df_map_merged["당해년도"] - df_map_merged["전년도"]) / df_map_merged["전년도"]) * 100, 0)
                    alarm_df = df_map_merged[df_map_merged["증감률(%)"] <= -5].copy()
                    
                    if alarm_df.empty:
                        st.success(f"✅ 선택한 기간 내 YoY 5% 이상 하락한 {map_usage} 리스크 업체가 없습니다.")
                    else:
                        st.warning(f"🚨 총 **{len(alarm_df)}**개의 {map_usage} 업체에서 5% 이상 하락 신호가 감지되었습니다. (지도에는 하락폭이 큰 주요 100개 업체를 표시합니다.)")
                        
                        alarm_df["감소량"] = alarm_df["전년도"] - alarm_df["당해년도"]
                        alarm_df = alarm_df.sort_values(by="감소량", ascending=False).head(100).reset_index(drop=True)
                        
                        alarm_df["증감"] = alarm_df["당해년도"] - alarm_df["전년도"]
                        
                        lats, lons, tooltips, colors, radiuses = [], [], [], [], []
                        for _, row in alarm_df.iterrows():
                            lat, lon = get_coord_from_df(row['도로명주소'], coord_df)
                            lats.append(lat)
                            lons.append(lon)
                            
                            rate = row['증감률(%)']
                            
                            if map_usage == "산업용":
                                if rate <= -20:
                                    level = "심각"
                                    colors.append([180, 0, 0, 255]) 
                                    radiuses.append(150)
                                elif rate <= -10:
                                    level = "경계"
                                    colors.append([255, 80, 80, 200]) 
                                    radiuses.append(100)
                                else:
                                    level = "주의"
                                    colors.append([255, 150, 150, 200]) 
                                    radiuses.append(80)
                            else: 
                                if rate <= -20:
                                    level = "심각"
                                    colors.append([0, 0, 180, 255]) 
                                    radiuses.append(150)
                                elif rate <= -10:
                                    level = "경계"
                                    colors.append([80, 150, 255, 200]) 
                                    radiuses.append(100)
                                else:
                                    level = "주의"
                                    colors.append([120, 180, 255, 200]) 
                                    radiuses.append(80)
                            
                            info = f"<b>{row['용도_태그']} {row['고객명']} <span style='color:red;'>[{level}]</span></b><br/>"
                            info += f"전년: {row['전년도']:,.0f} / 당해: {row['당해년도']:,.0f}<br/>"
                            info += f"증감률: <span style='color:red; font-weight:bold;'>{row['증감률(%)']:.1f}%</span><br/>"
                            info += f"<span style='font-size:0.8em; color:gray;'>{row['도로명주소']}</span>"
                            tooltips.append(info)
                            
                        alarm_df['lat'] = lats
                        alarm_df['lon'] = lons
                        alarm_df['tooltip'] = tooltips
                        alarm_df['color'] = colors
                        alarm_df['radius'] = radiuses
                        alarm_df = alarm_df.dropna(subset=['lat', 'lon'])
                        
                        if not alarm_df.empty:
                            layer = pdk.Layer(
                                "ScatterplotLayer",
                                data=alarm_df,
                                get_position='[lon, lat]',
                                get_color='color',     
                                get_radius='radius',   
                                pickable=True,
                                opacity=0.6,
                                filled=True,
                                stroked=True,
                                get_line_color=[255, 255, 255, 200],
                                line_width_min_pixels=1,
                                radius_max_pixels=40
                            )
                            
                            view_state = pdk.ViewState(
                                latitude=alarm_df['lat'].mean(),
                                longitude=alarm_df['lon'].mean(),
                                zoom=11,
                                pitch=40,
                            )
                            
                            r = pdk.Deck(
                                map_style=deck_map_style, 
                                layers=[layer],
                                initial_view_state=view_state,
                                tooltip={"html": "{tooltip}", "style": {"backgroundColor": "white", "color": "black", "font-family": "NanumGothic"}}
                            )
                            st.pydeck_chart(r)
                            
                            st.markdown(f"<br><b>📋 지도 표기 업체 요약표</b> <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                            
                            show_cols = ['용도_태그', '고객명', '도로명주소', '전년도', '당해년도', '증감', '증감률(%)']
                            df_show = alarm_df[show_cols].copy()
                            
                            df_show.insert(0, "No.", range(1, len(df_show) + 1))
                            df_show["비고"] = np.where(df_show["증감률(%)"] <= -99.9, "폐업의심", "")
                            
                            sum_prev_all = df_show["전년도"].sum()
                            sum_curr_all = df_show["당해년도"].sum()
                            sum_rate_all = ((sum_curr_all - sum_prev_all) / sum_prev_all * 100) if sum_prev_all > 0 else 0
                            
                            total_row = pd.DataFrame([{
                                "No.": "",
                                "용도_태그": "💡 총계",
                                "고객명": "",
                                "도로명주소": "",
                                "전년도": sum_prev_all,
                                "당해년도": sum_curr_all,
                                "증감": sum_curr_all - sum_prev_all,
                                "증감률(%)": sum_rate_all,
                                "비고": ""
                            }])
                            df_show = pd.concat([df_show, total_row], ignore_index=True)
                            
                            # 🟢 [수정됨] 1번째 사진 뒷 배경색 제거 요청 반영
                            # 글자 색상과 굵기는 유지하되 background-color만 제거 (총계는 회색 유지)
                            def highlight_map_summary_text_only(s):
                                # 총계 행은 기존처럼 회색 배경 유지
                                if s['용도_태그'] == "💡 총계":
                                    return ['background-color: #e0e2e6; font-weight: bold;'] * len(s)
                                
                                try:
                                    drop_val = float(s['증감률(%)'])
                                except:
                                    drop_val = 0
                                
                                # 배경색(background-color)은 빼고 글자색(color)과 굵기(font-weight)만 지정
                                if drop_val <= -15.0:
                                    return ['color: #b71c1c; font-weight: bold;'] * len(s) # 진한 빨강 글씨
                                elif drop_val <= -10.0:
                                    return ['color: #e65100; font-weight: bold;'] * len(s) # 주황색 글씨
                                return [''] * len(s)
                                
                            st.dataframe(center_style(df_show.style.format({"전년도": "{:,.0f}", "당해년도": "{:,.0f}", "증감": "{:,.0f}", "증감률(%)": "{:,.1f}"}).apply(highlight_map_summary_text_only, axis=1)), use_container_width=True, hide_index=True)
                        else:
                            st.error("매핑된 위경도 좌표가 없어 지도를 표시할 수 없습니다.")
                else:
                    st.info("비교할 과거 또는 당해 연도 데이터가 없습니다.")
        else:
            st.info("데이터에 '도로명주소', '고객명', '용도' 컬럼이 없거나 데이터가 부족하여 지도를 생성할 수 없습니다.")

        st.markdown("<hr style='border-top: 2px solid #1e3a8a; margin: 50px 0 20px 0;'>", unsafe_allow_html=True)

        # ─────────────────────────────────────────────────────────
        # 통합 분석 함수
        # ─────────────────────────────────────────────────────────
        def render_full_usage_report(usage_name, section_num, key_sfx, db_key):
            st.markdown(f"""<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 10px;"><h4 style="margin: 0;">📈 {section_num}. 용도별 판매량 분석 : {usage_name}</h4></div>""", unsafe_allow_html=True)
            
            if not df_long_rpt.empty:
                df_u = df_long_rpt[(df_long_rpt["그룹"] == usage_name) & (df_long_rpt["월"] <= max_month)]
                p_curr_act = df_u[(df_u["연"] == sel_year_rpt) & (df_u["계획/실적"] == "실적")].groupby("월")["값"].sum()
                p_prev_act = df_u[(df_u["연"] == sel_year_rpt-1) & (df_u["계획/실적"] == "실적")].groupby("월")["값"].sum()
            else:
                if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                    if "상품명" in df_csv_tab.columns:
                        csv_products = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                    else:
                        csv_products = pd.Series([""] * len(df_csv_tab))
                        
                    if usage_name == "산업용":
                        df_u_csv = df_csv_tab[(csv_products == "산업용") & (df_csv_tab["월_csv"] <= max_month)].copy()
                    else:
                        valid_biz_nospaces = ["냉난방용(업무)", "업무난방용", "주한미군"]
                        df_u_csv = df_csv_tab[(csv_products.isin(valid_biz_nospaces)) & (df_csv_tab["월_csv"] <= max_month)].copy()
                    
                    p_curr_act = df_u_csv[df_u_csv["연_csv"] == sel_year_rpt].groupby("월_csv")[val_col].sum()
                    p_prev_act = df_u_csv[df_u_csv["연_csv"] == sel_year_rpt-1].groupby("월_csv")[val_col].sum()
                else:
                    p_curr_act, p_prev_act = pd.Series(dtype=float), pd.Series(dtype=float)
            
            if agg_mode == "누적 실적 (1월~당월)":
                sum_act = p_curr_act.sum()
                sum_prev = p_prev_act.sum()
                top_title = f"**■ 누적 실적 비교 ({max_month}월 누적)**"
            else:
                sum_act = p_curr_act.get(max_month, 0)
                sum_prev = p_prev_act.get(max_month, 0)
                top_title = f"**■ 당월 실적 비교 ({max_month}월 당월)**"
            
            diff_prev = sum_act - sum_prev
            rate_prev = (sum_act / sum_prev * 100) if sum_prev > 0 else 0
            sign_prev = "+" if diff_prev > 0 else ""
            months_list = list(range(1, max_month + 1))
            
            desc_status = "감소" if diff_prev < 0 else "증가"
            st.markdown(
                f"""
                <div style="background-color: #f8f9fa; border-left: 5px solid #1e3a8a; padding: 15px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 15px; color: #1e3a8a; font-weight: 700; line-height: 1.6;">
                        💡 [요약] 당해 실적: {sum_act:,.0f} {unit_str}<br>
                        전년대비 <span style="color: {'#d32f2f' if diff_prev < 0 else '#1f77b4'};">{abs(diff_prev):,.0f} {unit_str} {desc_status} ({sign_prev}{rate_prev:.1f}%)</span>
                    </div>
                </div>
                """, unsafe_allow_html=True
            )

            # 그래프 데이터 준비
            vals_act = [p_curr_act.get(m, 0) for m in months_list]
            vals_prev = [p_prev_act.get(m, 0) for m in months_list]

            # 🟢 [수정됨] 2번째 사진 세로 눈금 동기화 요청 반영
            # 왼쪽 그래프(fig_c)와 오른쪽 그래프(fig_m)에 나타나는 모든 값 중 최댓값을 구해서 Y축 범위 통일
            graph_max_c = max([sum_prev, sum_act]) if months_list else 0
            graph_max_m = max(max(vals_act) if vals_act else 0, max(vals_prev) if vals_prev else 0)
            overall_max = max(graph_max_c, graph_max_m)
            
            # 여유 공간 10% 추가
            yaxis_range = [0, overall_max * 1.1 if overall_max > 0 else 100]
            
            col_c, col_m = st.columns([1, 2.5])
            with col_c:
                st.markdown(top_title + f" <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                fig_c = go.Figure()
                fig_c.update_layout(margin=dict(t=30, b=20, l=40, r=10), height=420, showlegend=False)
                # 동일한 Y축 범위 적용
                fig_c.update_yaxes(range=yaxis_range)
                fig_c.add_trace(go.Bar(x=[f"{sel_year_rpt-1}년<br>실적", f"{sel_year_rpt}년<br>실적"], y=[sum_prev, sum_act], marker_color=[COLOR_PREV, COLOR_ACT], text=[f"{sum_prev:,.0f}", f"{sum_act:,.0f}"], textposition='auto', textfont=dict(size=14)))
                st.plotly_chart(fig_c, use_container_width=True)
                
            with col_m:
                st.markdown(f"**■ 월별 실적 추이 (YoY)** <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                fig_m = go.Figure()
                fig_m.update_layout(barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1), margin=dict(t=30, b=20, l=40, r=10), height=420, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                # 동일한 Y축 범위 적용
                fig_m.update_yaxes(range=yaxis_range)
                
                fig_m.add_trace(go.Bar(x=months_list, y=vals_prev, name=f'{sel_year_rpt-1}년 실적', marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v>0 else "" for v in vals_prev], textposition='auto', textfont=dict(size=11)))
                fig_m.add_trace(go.Bar(x=months_list, y=vals_act, name=f'{sel_year_rpt}년 실적', marker_color=COLOR_ACT, text=[f"{v:,.0f}" if v>0 else "" for v in vals_act], textposition='auto', textfont=dict(size=11)))
                st.plotly_chart(fig_m, use_container_width=True)

            render_comment_section(f"📝 {usage_name} 세부 코멘트 작성", db_key, curr_db, comments_db, 100, f"{usage_name}의 월별 편차 원인 및 특이사항을 기록하세요.", f"{usage_name}_{key_sfx}")
            st.markdown("<hr style='border-top: 1px dashed #ccc; margin: 30px 0;'>", unsafe_allow_html=True)

            if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                if "상품명" in df_csv_tab.columns:
                    csv_products = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                else:
                    csv_products = pd.Series([""] * len(df_csv_tab))
                
                if agg_mode == "누적 실적 (1월~당월)":
                    month_mask = (df_csv_tab["월_csv"] <= max_month)
                else:
                    month_mask = (df_csv_tab["월_csv"] == max_month)

                if usage_name == "산업용":
                    df_sub_filtered = df_csv_tab[(csv_products == "산업용") & month_mask].copy()
                else: 
                    valid_biz_nospaces = ["냉난방용(업무)", "업무난방용", "주한미군"]
                    df_sub_filtered = df_csv_tab[(csv_products.isin(valid_biz_nospaces)) & month_mask].copy()

                st.markdown(f"**🔍 {usage_name} 개별 고객 상세 차트** <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                
                if not df_sub_filtered.empty and "고객명" in df_sub_filtered.columns:
                    top_cust_data = df_sub_filtered[df_sub_filtered["연_csv"] == sel_year_rpt].groupby("고객명")[val_col].sum().sort_values(ascending=False).head(50)
                    top_customers = top_cust_data.index.tolist()
                    sel_cust = st.selectbox(f"상세 분석할 고객명을 선택하세요 ({usage_name})", ["선택 안함"] + top_customers, key=f"sel_cust_{usage_name}_{key_sfx}")

                    if sel_cust != "선택 안함":
                        c_data = df_csv_tab[df_csv_tab["고객명"] == sel_cust]
                        c_grp = c_data.groupby(["연_csv", "월_csv"], as_index=False)[val_col].sum()
                        
                        y_cur = c_grp[(c_grp["연_csv"] == sel_year_rpt) & (c_grp["월_csv"] <= max_month)]
                        y_prev = c_grp[(c_grp["연_csv"] == sel_year_rpt - 1) & (c_grp["월_csv"] <= max_month)]
                        
                        if agg_mode == "누적 실적 (1월~당월)":
                            sum_cur_c = y_cur[val_col].sum()
                            sum_prev_c = y_prev[val_col].sum()
                            chart_title = f"'{sel_cust}' 누적 사용량 ({max_month}월 누적)"
                        else:
                            sum_cur_c = y_cur[y_cur["월_csv"] == max_month][val_col].sum()
                            sum_prev_c = y_prev[y_prev["월_csv"] == max_month][val_col].sum()
                            chart_title = f"'{sel_cust}' 당월 사용량 ({max_month}월 당월)"
                            
                        diff_val = sum_cur_c - sum_prev_c
                        rate_val = (sum_cur_c / sum_prev_c * 100) if sum_prev_c > 0 else 0
                        sign_str = "+" if diff_val > 0 else ""
                        yoy_text = f"전년대비 증감: {sign_str}{diff_val:,.0f} ({rate_val:.1f}%)"
                        
                        cc1, cc2 = st.columns([1, 2])
                        with cc1:
                            fig_cust_cum = go.Figure()
                            fig_cust_cum.update_layout(title=chart_title, margin=dict(t=50, b=20, l=40, r=10), height=350)
                            fig_cust_cum.add_trace(go.Bar(x=[f"{sel_year_rpt-1}년", f"{sel_year_rpt}년"], y=[sum_prev_c, sum_cur_c], marker_color=[COLOR_PREV, COLOR_ACT], text=[f"{sum_prev_c:,.0f}", f"{sum_cur_c:,.0f}"], textposition='auto'))
                            fig_cust_cum.add_annotation(x=0.5, y=1.05, xref="paper", yref="paper", text=f"<b>{yoy_text}</b>", showarrow=False, font=dict(size=13, color="#d32f2f" if diff_val < 0 else "#1f77b4"), bgcolor="#f8f9fa", bordercolor="#d0d7e5", borderwidth=1, borderpad=4)
                            st.plotly_chart(fig_cust_cum, use_container_width=True)
                            
                        with cc2:
                            fig_cust_mon = go.Figure()
                            fig_cust_mon.update_layout(title=f"'{sel_cust}' 월별 사용량 추이", barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1), margin=dict(t=50, b=20, l=40, r=10), height=350, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                            months_c = list(range(1, max_month + 1))
                            cur_vals = [y_cur[y_cur['월_csv']==m][val_col].sum() for m in months_c]
                            prev_vals = [y_prev[y_prev['월_csv']==m][val_col].sum() for m in months_c]
                            
                            fig_cust_mon.add_trace(go.Bar(x=months_c, y=prev_vals, name=f"{sel_year_rpt-1}년", marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v>0 else "" for v in prev_vals], textposition='auto', textfont=dict(size=11)))
                            fig_cust_mon.add_trace(go.Bar(x=months_c, y=cur_vals, name=f"{sel_year_rpt}년", marker_color=COLOR_ACT, text=[f"{v:,.0f}" if v>0 else "" for v in cur_vals], textposition='auto', textfont=dict(size=11)))
                            st.plotly_chart(fig_cust_mon, use_container_width=True)

        render_full_usage_report("산업용", "2", key_sfx, "ind")
        st.markdown("<hr style='margin: 50px 0; border-top: 2px solid #ccc;'>", unsafe_allow_html=True)
        render_full_usage_report("업무용", "3", key_sfx, "biz")

        # ─────────────────────────────────────────────────────────
        # 4. 보고서 출력
        # ─────────────────────────────────────────────────────────
        st.markdown("<hr style='border-top: 2px solid #bbb; margin: 40px 0 20px 0;'>", unsafe_allow_html=True)
        st.markdown("### 🖨️ 4. 보고서 출력")
        
        st.markdown("""
            <style>
            @media print {
                header[data-testid="stHeader"] { display: none !important; }
                section[data-testid="stSidebar"] { display: none !important; }
                div[data-testid="stToolbar"] { display: none !important; }
                iframe[title="st.iframe"] { display: none !important; }
            }
            </style>
        """, unsafe_allow_html=True)
        
        st.components.v1.html("""
            <button onclick="window.parent.print()" style="padding: 12px 20px; font-size: 16px; border-radius: 8px; background-color: #1e3a8a; color: white; border: none; cursor: pointer; width: 100%; font-weight: bold; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin: 2px;">
                🖨️ 현재 화면 전체를 PDF로 다운로드 (인쇄)
            </button>
        """, height=70)
