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

try:
    from github import Github
    HAS_GITHUB = True
except ImportError:
    HAS_GITHUB = False


def set_korean_font():
    try:
        ttf = Path(__file__).parent / "NanumGothic-Regular.ttf"
        if ttf.exists():
            mpl.font_manager.fontManager.addfont(str(ttf))
            mpl.rcParams["font.family"] = "NanumGothic"
            mpl.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass

set_korean_font()
st.set_page_config(page_title="대용량 수요처 이상 감지 대시보드", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"
REPO_NAME = "Han11112222/quarterly-sales-report"
GITHUB_RAW_BASE = f"https://raw.githubusercontent.com/{REPO_NAME}/main"
COMMENT_DB_FILE = "report_comments_db.json"

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


# ─────────────────────────────────────────────────────────
# 코멘트 DB
# ─────────────────────────────────────────────────────────
def load_comments_db():
    if os.path.exists(COMMENT_DB_FILE):
        try:
            with open(COMMENT_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_comments_db(db_data):
    try:
        with open(COMMENT_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db_data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass
    if not HAS_GITHUB:
        return
    try:
        token = st.secrets["GITHUB_TOKEN"]
        if token:
            g = Github(token)
            repo = g.get_repo(REPO_NAME)
            content_string = json.dumps(db_data, ensure_ascii=False, indent=4)
            try:
                contents = repo.get_contents(COMMENT_DB_FILE)
                repo.update_file(contents.path, "Update comments", content_string, contents.sha)
            except Exception:
                repo.create_file(COMMENT_DB_FILE, "Create comments db", content_string)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────────────────
def clean_korean_finance_number(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(",", "").strip()
    if not s: return 0.0
    if s.endswith("-"): s = "-" + s[:-1]
    elif s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    s = re.sub(r"[^\d\.-]", "", s)
    try: return float(s)
    except: return 0.0

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

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns: out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

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
    try:
        xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
        needed = ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]
        out: Dict[str, pd.DataFrame] = {}
        for name in needed:
            if name in xls.sheet_names: out[name] = xls.parse(name)
        return out
    except Exception as e:
        st.warning(f"엑셀 파일 로드 오류: {e}")
        return {}

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    long_dict: Dict[str, pd.DataFrame] = {}
    if ("계획_부피" in sheets) and ("실적_부피" in sheets):
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if ("계획_열량" in sheets) and ("실적_열량" in sheets):
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

def load_safe_csv(file_bytes) -> pd.DataFrame:
    for enc in ["utf-8-sig", "cp949", "utf-8", "euc-kr"]:
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
                except Exception:
                    pass
    lat = 35.8714 + random.uniform(-0.06, 0.06)
    lon = 128.6014 + random.uniform(-0.06, 0.06)
    return lat, lon

KAKAO_REST_API_KEY = "d9532fed7e56e09fe392c3482b915a20"

def get_kakao_route(start_lon, start_lat, end_lon, end_lat):
    url = "https://apis-navi.kakaomobility.com/v1/directions"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {
        "origin": f"{start_lon},{start_lat}",
        "destination": f"{end_lon},{end_lat}",
        "priority": "RECOMMEND",
    }
    try:
        response = requests.get(url, headers=headers, params=params, timeout=5)
        if response.status_code == 200:
            res_json = response.json()
            path_coords = []
            if res_json.get("routes"):
                sections = res_json["routes"][0]["sections"]
                for section in sections:
                    for road in section["roads"]:
                        vertices = road["vertexes"]
                        for i in range(0, len(vertices), 2):
                            path_coords.append([vertices[i], vertices[i + 1]])
            return path_coords
        return None
    except Exception:
        return None


# ─────────────────────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────────────────────
st.title("📊 대용량 수요처 이상 감지 대시보드")

with st.sidebar:
    page_sel = st.radio(
        "📌 메뉴",
        ["1. 대용량 수요처 이상 감지", "2. 개별 고객 검색"],
        index=0, key="page_sel"
    )
    st.markdown("---")

with st.sidebar:
    st.header("📂 데이터 및 설정")

    st.subheader("1. 판매량 데이터 (요약/엑셀)")
    src_sales = st.radio("판매량 데이터 소스", ["레포 파일 사용", "엑셀 업로드(.xlsx)"], index=0, key="rpt_sales_src")
    excel_bytes = None
    if src_sales == "엑셀 업로드(.xlsx)":
        up_sales = st.file_uploader("판매량(계획_실적).xlsx 형식", type=["xlsx"], key="rpt_sales_uploader")
        if up_sales is not None: excel_bytes = up_sales.getvalue()
    else:
        try:
            path_sales = Path(__file__).parent / DEFAULT_SALES_XLSX
            if path_sales.exists(): excel_bytes = path_sales.read_bytes()
        except Exception:
            pass

    st.markdown("---")
    st.subheader("2. 업종별 데이터 (상세/CSV)")
    src_csv = st.radio("업종별 데이터 소스", ["레포 파일 사용", "CSV 업로드(.csv)"], index=0, key="csv_src")
    if src_csv == "CSV 업로드(.csv)":
        up_csvs = st.file_uploader("가정용외_*.csv 형식 (다중 업로드 가능)", type=["csv"],
                                    accept_multiple_files=True, key="csv_uploader")
        if up_csvs:
            df_list = []
            for f in up_csvs:
                df_tmp = load_safe_csv(f.getvalue())
                if not df_tmp.empty:
                    df_list.append(df_tmp)
            if df_list: st.session_state['merged_csv_df'] = pd.concat(df_list, ignore_index=True)
        else:
            if 'merged_csv_df' in st.session_state: del st.session_state['merged_csv_df']

    st.markdown("---")
    st.subheader("🗺️ 3. 지도 위경도 데이터 (CSV)")
    src_coord = st.radio("위경도 데이터 소스", ["레포 파일(깃허브) 사용", "CSV 업로드(.csv)"], index=0, key="coord_src")
    coord_df = pd.DataFrame()
    if src_coord == "CSV 업로드(.csv)":
        up_coord = st.file_uploader("위경도 매핑 파일 업로드", type=["csv"], key="coord_uploader")
        if up_coord:
            coord_df = load_safe_csv(up_coord.getvalue())
    else:
        try:
            coord_path = Path(__file__).parent / "address_with_latlon.csv"
            if coord_path.exists():
                coord_df = load_safe_csv(coord_path.read_bytes())
            else:
                res = requests.get(f"{GITHUB_RAW_BASE}/address_with_latlon.csv", timeout=5)
                if res.status_code == 200:
                    coord_df = load_safe_csv(res.content)
        except Exception:
            pass


# ─────────────────────────────────────────────────────────
# 본문
# ─────────────────────────────────────────────────────────
long_dict_rpt: Dict[str, pd.DataFrame] = {}
if excel_bytes is not None:
    sheets_rpt = load_all_sheets(excel_bytes)
    long_dict_rpt = build_long_dict(sheets_rpt)

df_csv = pd.DataFrame()
if src_csv == "레포 파일 사용":
    try:
        repo_dir = Path(__file__).parent
        all_csvs = list(repo_dir.glob("*가정용외*.csv")) + list(repo_dir.glob("가정용외*.csv"))
        all_csvs = list(set(all_csvs))
        csv_list = []
        for p in all_csvs:
            try:
                temp_df = pd.read_csv(p, encoding="utf-8-sig", thousands=',')
                temp_df.columns = temp_df.columns.str.strip()
                csv_list.append(temp_df)
            except Exception:
                try:
                    temp_df = pd.read_csv(p, encoding="cp949", thousands=',')
                    temp_df.columns = temp_df.columns.str.strip()
                    csv_list.append(temp_df)
                except Exception:
                    pass
        if csv_list: df_csv = pd.concat(csv_list, ignore_index=True)
    except Exception:
        pass

if df_csv.empty and 'merged_csv_df' in st.session_state:
    df_csv = st.session_state['merged_csv_df'].copy()

if not df_csv.empty:
    if "사용량(mj)" in df_csv.columns: df_csv["사용량(mj)"] = df_csv["사용량(mj)"].apply(clean_korean_finance_number)
    if "사용량(m3)" in df_csv.columns: df_csv["사용량(m3)"] = df_csv["사용량(m3)"].apply(clean_korean_finance_number)

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
                max_month_val = actual_data[actual_data["연"] == max_year]["월"].max()
                default_y_index = years_available.index(max_year) if max_year in years_available else len(years_available) - 1
                default_m_index = int(max_month_val - 1)

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

            csv_max_date = df_csv_tab["날짜_파싱"].max()
            if pd.notna(csv_max_date):
                if csv_max_date.year in years_available:
                    default_y_index = years_available.index(csv_max_date.year)
                default_m_index = int(csv_max_date.month - 1)

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

        st.markdown("<hr style='margin: 10px 0 30px 0;'>", unsafe_allow_html=True)

        if page_sel == "2. 개별 고객 검색":
            def cs_upjong_map(df_part, names, is_biz):
                """고객별 업종. 산업용='업종', 업무용='업종분류' (비어 있으면 서로 보완)
                names: 고객명 Index, is_biz: names와 같은 인덱스의 bool Series"""
                def _first(col):
                    if col in df_part.columns:
                        s = df_part.groupby("고객명")[col].first()
                    else:
                        s = pd.Series(dtype=object)
                    return s.reindex(names).fillna("").astype(str).str.strip()
                up_a = _first("업종")
                up_b = _first("업종분류")
                up = up_a.where(up_a != "", up_b)
                up = up.where(~(is_biz & (up_b != "")), up_b)
                return up

            def cs_show_df(df, row_style_fn, fmt, key, height=None):
                """행 단위 배경 하이라이트를 적용한 HTML 표 출력
                (st.dataframe/Styler를 쓰지 않으므로 브라우저 렌더링 문제가 없음)"""
                import html as _html
                th_css = ("padding:6px 10px; text-align:left; background-color:#f1f3f5; color:#212529; "
                          "border-bottom:2px solid #ced4da; position:sticky; top:0; white-space:nowrap;")
                head = "".join(f"<th style='{th_css}'>{_html.escape(str(c))}</th>" for c in df.columns)
                num_cols = {c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])}
                body = []
                for _, row in df.iterrows():
                    css = row_style_fn(row)[0]
                    tds = []
                    for c in df.columns:
                        v = row[c]
                        if pd.isna(v):
                            txt = ""
                        elif c in fmt:
                            txt = fmt[c].format(v)
                        else:
                            txt = str(v)
                        td_css = ("padding:5px 10px; border-bottom:1px solid #e9ecef; white-space:nowrap; "
                                  f"text-align:{'right' if c in num_cols else 'left'};")
                        tds.append(f"<td style='{td_css}'>{_html.escape(txt)}</td>")
                    body.append(f"<tr style='background-color:#ffffff; color:#212529; {css}'>{''.join(tds)}</tr>")
                max_h = f"max-height:{height}px;" if height else ""
                st.markdown(
                    f"<div style='overflow:auto; {max_h} border:1px solid #dee2e6; border-radius:6px;'>"
                    f"<table style='border-collapse:collapse; width:100%; font-size:13px;'>"
                    f"<thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table></div>",
                    unsafe_allow_html=True
                )

            # ═════════════════════════════════════════════════════════
            # [신규 추가] 고객 검색 (검색 → 지도 표시 → 개별 분석)
            # ※ 모든 변수는 cs_ 접두어를 사용하므로 기존 코드와 충돌하지 않습니다.
            # ═════════════════════════════════════════════════════════
            st.markdown(
                f"### 🔎 개별 고객 검색 "
                f"<span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>",
                unsafe_allow_html=True
            )
            st.caption("※ 고객명을 검색하면 해당 업체가 지도에 표시되고, 바로 아래에서 사용량 분석을 확인할 수 있습니다. (상단 기준 일자·집계 기준이 적용됩니다)")

            cs_found_names = []      # 검색에 걸린 고객명 (아래 랭킹 영역에서 사용)
            cs_sel_name = None       # 선택된 고객명
            if (not df_csv_tab.empty and "고객명" in df_csv_tab.columns and val_col in df_csv_tab.columns):
                cs_c1, cs_c2, cs_c3 = st.columns([2, 1, 1])
                with cs_c1:
                    cs_query = st.text_input(
                        "고객명 입력 (일부만 입력해도 검색됩니다)",
                        placeholder="예: 대성, 한국 ...",
                        key=f"cs_query{key_sfx}"
                    )
                with cs_c2:
                    cs_comp_mode = st.radio(
                        "비교 기준", ["YoY", "전월대비"], index=0, horizontal=True,
                        key=f"cs_comp{key_sfx}"
                    )
                with cs_c3:
                    cs_map_style_ui = st.radio(
                        "지도 배경 테마", ["다크 모드 (기본)", "일반 도로 지도"], index=0, horizontal=True,
                        key=f"cs_map_style{key_sfx}"
                    )
                cs_map_style = (
                    "https://basemaps.cartocdn.com/gl/dark-matter-gl-style/style.json"
                    if cs_map_style_ui == "다크 모드 (기본)"
                    else "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json"
                )

                cs_year = sel_year_rpt
                cs_month = max_month
                if cs_comp_mode == "YoY":
                    cs_prev_year, cs_prev_month = cs_year - 1, cs_month
                    cs_prev_col, cs_cur_col, cs_diff_label = "전년도", "당해년도", "전년대비"
                else:
                    cs_prev_year = cs_year if cs_month > 1 else cs_year - 1
                    cs_prev_month = cs_month - 1 if cs_month > 1 else 12
                    cs_prev_col, cs_cur_col, cs_diff_label = "전월", "당월", "전월대비"

                def cs_mask(df, y, m):
                    if agg_mode == "누적 실적 (1월~당월)":
                        return (df["연_csv"] == y) & (df["월_csv"] <= m)
                    return (df["연_csv"] == y) & (df["월_csv"] == m)

                def cs_find_coord(addr):
                    """주소 → (위도, 경도). 매핑되지 않으면 (None, None)"""
                    if pd.isna(addr) or not str(addr).strip():
                        return None, None
                    if coord_df.empty or len(coord_df.columns) < 3:
                        return None, None
                    key = re.sub(r'\(.*?\)', '', str(addr)).split(',')[0].strip().replace(" ", "")
                    if not key:
                        return None, None
                    col_addr = coord_df[coord_df.columns[0]].astype(str).str.replace(" ", "", regex=False)
                    hit = coord_df[col_addr.str.contains(re.escape(key), na=False)]
                    if hit.empty:
                        return None, None
                    try:
                        return float(hit.iloc[0][coord_df.columns[1]]), float(hit.iloc[0][coord_df.columns[2]])
                    except Exception:
                        return None, None

                cs_q = cs_query.strip().replace(" ", "")
                if not cs_q:
                    st.info("🔍 검색어를 입력하면 해당 고객이 지도에 표시되고 분석 차트가 나타납니다.")
                else:
                    cs_name_ns = df_csv_tab["고객명"].fillna("").astype(str).str.replace(" ", "", regex=False)
                    cs_hit = df_csv_tab[cs_name_ns.str.contains(re.escape(cs_q), case=False, na=False)]

                    if cs_hit.empty:
                        st.warning(f"'{cs_query}' 에 해당하는 고객이 없습니다.")
                    else:
                        # ── 검색 결과 요약표 만들기 ──
                        cs_cur = cs_hit[cs_mask(cs_hit, cs_year, cs_month)].groupby("고객명")[val_col].sum()
                        cs_pre = cs_hit[cs_mask(cs_hit, cs_prev_year, cs_prev_month)].groupby("고객명")[val_col].sum()
                        cs_info = cs_hit.drop_duplicates("고객명").set_index("고객명")

                        cs_tbl = pd.DataFrame(index=cs_info.index)
                        cs_tbl["용도"] = cs_info["용도"].fillna("") if "용도" in cs_info.columns else ""
                        # 업종: 산업용은 '업종', 업무용은 '업종분류' 사용 (비어 있으면 다른 컬럼으로 보완)
                        cs_tbl["업종"] = cs_upjong_map(cs_hit, cs_tbl.index, cs_tbl["용도"].astype(str) == "업무용")
                        cs_tbl["도로명주소"] = cs_info["도로명주소"].fillna("") if "도로명주소" in cs_info.columns else ""
                        cs_tbl[cs_prev_col] = cs_pre.reindex(cs_tbl.index).fillna(0)
                        cs_tbl[cs_cur_col] = cs_cur.reindex(cs_tbl.index).fillna(0)
                        cs_tbl["증감"] = cs_tbl[cs_cur_col] - cs_tbl[cs_prev_col]
                        cs_tbl["증감률(%)"] = np.where(
                            cs_tbl[cs_prev_col] > 0,
                            cs_tbl["증감"] / cs_tbl[cs_prev_col] * 100,
                            0
                        )
                        cs_tbl = cs_tbl.sort_values([cs_cur_col, cs_prev_col], ascending=False)

                        cs_total_hit = len(cs_tbl)
                        if cs_total_hit > 100:
                            cs_tbl = cs_tbl.head(100)
                            st.caption(f"※ 검색 결과 {cs_total_hit}개 중 사용량 상위 100개만 표시합니다. 검색어를 더 구체적으로 입력해 보세요.")

                        st.success(f"✅ '{cs_query}' 검색 결과: {len(cs_tbl)}개 업체")

                        cs_names = list(cs_tbl.index)
                        cs_sel = st.selectbox(
                            "분석할 고객 선택",
                            cs_names,
                            key=f"cs_sel{key_sfx}"
                        )
                        cs_found_names = list(cs_tbl.index)
                        cs_sel_name = cs_sel

                        # ── 지도 표시 ──
                        cs_rows, cs_unmapped = [], []
                        for cs_name, cs_r in cs_tbl.iterrows():
                            cs_lat, cs_lon = cs_find_coord(cs_r["도로명주소"])
                            if cs_lat is None:
                                cs_unmapped.append(cs_name)
                                continue
                            cs_is_sel = (cs_name == cs_sel)
                            cs_usage = str(cs_r["용도"])
                            if cs_is_sel:
                                cs_color, cs_radius = [255, 140, 0, 255], 200
                            elif cs_usage == "산업용":
                                cs_color, cs_radius = [255, 80, 80, 200], 100
                            elif cs_usage == "업무용":
                                cs_color, cs_radius = [80, 150, 255, 200], 100
                            else:
                                cs_color, cs_radius = [120, 120, 120, 200], 100
                            cs_rate = cs_r["증감률(%)"]
                            cs_tip = (
                                f"<b>[{cs_usage}] {cs_name}</b><br/>"
                                f"{cs_prev_col}: {cs_r[cs_prev_col]:,.0f} / {cs_cur_col}: {cs_r[cs_cur_col]:,.0f}<br/>"
                                f"증감률: <span style='color:{'red' if cs_rate < 0 else 'blue'}; font-weight:bold;'>{cs_rate:.1f}%</span><br/>"
                                f"<span style='font-size:0.8em; color:gray;'>{cs_r['도로명주소']}</span>"
                            )
                            cs_rows.append({
                                "고객명": cs_name, "lat": cs_lat, "lon": cs_lon,
                                "color": cs_color, "radius": cs_radius, "tooltip": cs_tip,
                                "sel": cs_is_sel
                            })

                        if cs_rows:
                            cs_map_df = pd.DataFrame(cs_rows)
                            # 선택 고객이 맨 위에 그려지도록 정렬
                            cs_map_df = cs_map_df.sort_values("sel").reset_index(drop=True)
                            cs_sel_pt = cs_map_df[cs_map_df["sel"]]
                            if not cs_sel_pt.empty:
                                cs_center_lat = float(cs_sel_pt.iloc[0]["lat"])
                                cs_center_lon = float(cs_sel_pt.iloc[0]["lon"])
                                cs_zoom = 13
                            else:
                                cs_center_lat = float(cs_map_df["lat"].mean())
                                cs_center_lon = float(cs_map_df["lon"].mean())
                                cs_zoom = 11

                            cs_layer = pdk.Layer(
                                "ScatterplotLayer", data=cs_map_df,
                                get_position='[lon, lat]', get_color='color', get_radius='radius',
                                pickable=True, opacity=0.8, filled=True, stroked=True,
                                get_line_color=[255, 255, 255, 220], line_width_min_pixels=1,
                                radius_max_pixels=40
                            )
                            cs_deck = pdk.Deck(
                                map_style=cs_map_style,
                                layers=[cs_layer],
                                initial_view_state=pdk.ViewState(
                                    latitude=cs_center_lat, longitude=cs_center_lon,
                                    zoom=cs_zoom, pitch=0
                                ),
                                tooltip={"html": "{tooltip}", "style": {"backgroundColor": "white", "color": "black"}}
                            )
                            st.pydeck_chart(cs_deck)
                            st.caption("🟠 선택 고객 / 🔴 산업용 / 🔵 업무용 / ⚪ 기타 용도")
                        else:
                            st.error("검색된 업체의 위경도 좌표가 매핑되어 있지 않아 지도에 표시할 수 없습니다. (아래 분석은 정상 제공됩니다)")

                        if cs_unmapped and cs_rows:
                            st.caption(f"※ 위경도 미매핑으로 지도에서 제외된 업체: {len(cs_unmapped)}개")

                        # ── 검색 결과 표 ──
                        cs_show = cs_tbl.reset_index()
                        st.dataframe(
                            cs_show,
                            column_config={
                                cs_prev_col: st.column_config.NumberColumn(format="%.0f"),
                                cs_cur_col: st.column_config.NumberColumn(format="%.0f"),
                                "증감": st.column_config.NumberColumn(format="%.0f"),
                                "증감률(%)": st.column_config.NumberColumn(format="%.1f"),
                            },
                            hide_index=True,
                            use_container_width=True,
                            key=f"cs_table{key_sfx}"
                        )

                        # ── 선택 고객 개별 분석 ──
                        st.markdown(f"#### 📈 선택 고객 분석 : {cs_sel}")
                        cs_data = df_csv_tab[df_csv_tab["고객명"] == cs_sel]

                        cs_sum_cur = cs_data[cs_mask(cs_data, cs_year, cs_month)][val_col].sum()
                        cs_sum_prev = cs_data[cs_mask(cs_data, cs_prev_year, cs_prev_month)][val_col].sum()
                        cs_diff = cs_sum_cur - cs_sum_prev
                        cs_rate_all = (cs_diff / cs_sum_prev * 100) if cs_sum_prev > 0 else 0
                        cs_sign = "+" if cs_diff > 0 else ""

                        if cs_comp_mode == "YoY":
                            cs_prev_name, cs_cur_name = f"{cs_prev_year}년", f"{cs_year}년"
                            cs_prev_legend, cs_cur_legend = f"{cs_prev_year}년 실적", f"{cs_year}년 실적"
                        else:
                            cs_prev_name, cs_cur_name = f"전월({cs_prev_month}월)", f"당월({cs_month}월)"
                            cs_prev_legend, cs_cur_legend = "전월 실적", "당월 실적"

                        cs_period_title = (
                            f"'{cs_sel}' 누적 사용량 ({cs_month}월 누적)"
                            if agg_mode == "누적 실적 (1월~당월)"
                            else f"'{cs_sel}' 당월 사용량 ({cs_month}월)"
                        )

                        st.markdown(
                            f"""<div style="background-color: #f8f9fa; border-left: 5px solid #1e3a8a; padding: 15px;
                                margin-bottom: 20px; border-radius: 4px;">
                                <div style="font-size: 15px; color: #1e3a8a; font-weight: 700; line-height: 1.6;">
                                    💡 [요약] 당해 실적: {cs_sum_cur:,.0f} {unit_str}<br>
                                    {cs_diff_label} <span style="color: {'#d32f2f' if cs_diff < 0 else '#1f77b4'};">
                                    {abs(cs_diff):,.0f} {unit_str} {'감소' if cs_diff < 0 else '증가'} ({cs_sign}{cs_rate_all:.1f}%)</span>
                                </div></div>""",
                            unsafe_allow_html=True
                        )

                        cs_months = list(range(1, cs_month + 1))
                        cs_vals_cur = [
                            cs_data[(cs_data["연_csv"] == cs_year) & (cs_data["월_csv"] == m)][val_col].sum()
                            for m in cs_months
                        ]
                        if cs_comp_mode == "YoY":
                            cs_vals_prev = [
                                cs_data[(cs_data["연_csv"] == cs_year - 1) & (cs_data["월_csv"] == m)][val_col].sum()
                                for m in cs_months
                            ]
                        else:
                            cs_vals_prev = []
                            for m in cs_months:
                                if m > 1:
                                    cs_vals_prev.append(
                                        cs_data[(cs_data["연_csv"] == cs_year) & (cs_data["월_csv"] == m - 1)][val_col].sum())
                                else:
                                    cs_vals_prev.append(
                                        cs_data[(cs_data["연_csv"] == cs_year - 1) & (cs_data["월_csv"] == 12)][val_col].sum())

                        cs_a1, cs_a2 = st.columns([1, 2])
                        with cs_a1:
                            cs_fig1 = go.Figure()
                            cs_fig1.update_layout(title=cs_period_title, margin=dict(t=50, b=20, l=40, r=10), height=350)
                            cs_fig1.add_trace(go.Bar(
                                x=[cs_prev_name, cs_cur_name], y=[cs_sum_prev, cs_sum_cur],
                                marker_color=[COLOR_PREV, COLOR_ACT],
                                text=[f"{cs_sum_prev:,.0f}", f"{cs_sum_cur:,.0f}"],
                                textposition='auto'
                            ))
                            cs_fig1.add_annotation(
                                x=0.5, y=1.05, xref="paper", yref="paper",
                                text=f"<b>{cs_diff_label} 증감: {cs_sign}{cs_diff:,.0f} ({cs_rate_all:.1f}%)</b>",
                                showarrow=False,
                                font=dict(size=13, color="#d32f2f" if cs_diff < 0 else "#1f77b4"),
                                bgcolor="#f8f9fa", bordercolor="#d0d7e5", borderwidth=1, borderpad=4
                            )
                            st.plotly_chart(cs_fig1, key=f"cs_fig_cum{key_sfx}_{cs_sel}", use_container_width=True)

                        with cs_a2:
                            cs_fig2 = go.Figure()
                            cs_fig2.update_layout(
                                title=f"'{cs_sel}' 월별 사용량 추이", barmode='group',
                                xaxis=dict(tickmode='linear', tick0=1, dtick=1),
                                margin=dict(t=50, b=20, l=40, r=10), height=350,
                                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                            )
                            cs_fig2.add_trace(go.Bar(
                                x=cs_months, y=cs_vals_prev, name=cs_prev_legend, marker_color=COLOR_PREV,
                                text=[f"{v:,.0f}" if v > 0 else "" for v in cs_vals_prev],
                                textposition='auto', textfont=dict(size=11)))
                            cs_fig2.add_trace(go.Bar(
                                x=cs_months, y=cs_vals_cur, name=cs_cur_legend, marker_color=COLOR_ACT,
                                text=[f"{v:,.0f}" if v > 0 else "" for v in cs_vals_cur],
                                textposition='auto', textfont=dict(size=11)))
                            st.plotly_chart(cs_fig2, key=f"cs_fig_mon{key_sfx}_{cs_sel}", use_container_width=True)
            else:
                st.info("고객 검색을 위한 업종별 CSV 데이터('고객명' 컬럼)가 없습니다.")


            # ─────────────────────────────────────────────────────────
            # 누적 랭킹 (1월~기준월) : 산업용 Top30 / 업무용 Top30 + 검색 고객 순위
            # ─────────────────────────────────────────────────────────
            st.markdown("<hr style='border-top: 2px solid #1e3a8a; margin: 40px 0 30px 0;'>", unsafe_allow_html=True)
            st.markdown(
                f"### 🏆 누적 판매량 랭킹 (1~{max_month}월 누적) "
                f"<span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>",
                unsafe_allow_html=True
            )
            st.caption(
                f"※ {sel_year_rpt}년 1월~{max_month}월 누적 사용량 기준 순위이며 전년 동기 누적과 비교합니다. "
                f"(상단 '집계 기준' 선택과 무관)  |  검색 고객 표: 노란 배경 = Top30 업체  |  "
                f"랭킹 표: 연한 주황 = 검색된 고객, 진한 주황 = 선택 고객"
            )

            if (not df_csv_tab.empty and "고객명" in df_csv_tab.columns
                    and val_col in df_csv_tab.columns and "연_csv" in df_csv_tab.columns):
                rk_year, rk_month = sel_year_rpt, max_month
                cs_found_set = set(cs_found_names)
                CS_FMT = {"누적 판매량": "{:,.0f}", "전년 동기": "{:,.0f}", "증감": "{:+,.0f}",
                          "증감률(%)": "{:+.1f}", "구성비(%)": "{:.1f}"}

                def cs_usage_subset(df, usage):
                    prod = (df["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                            if "상품명" in df.columns else pd.Series("", index=df.index))
                    use = (df["용도"].astype(str).str.strip()
                           if "용도" in df.columns else pd.Series("", index=df.index))
                    if usage == "산업용":
                        return df[(use == "산업용") | (prod == "산업용")].copy()
                    return df[(use == "업무용") | prod.isin(["냉난방용(업무)", "업무난방용", "주한미군"])].copy()

                def cs_build_rank(usage):
                    sub = cs_usage_subset(df_csv_tab, usage)
                    if sub.empty:
                        return pd.DataFrame()
                    cur = sub[(sub["연_csv"] == rk_year) & (sub["월_csv"] <= rk_month)].groupby("고객명")[val_col].sum()
                    prv = sub[(sub["연_csv"] == rk_year - 1) & (sub["월_csv"] <= rk_month)].groupby("고객명")[val_col].sum()
                    names = cur.index.union(prv.index)
                    if len(names) == 0:
                        return pd.DataFrame()
                    rk = pd.DataFrame(index=names)
                    rk["누적 판매량"] = cur.reindex(names).fillna(0)
                    rk["전년 동기"] = prv.reindex(names).fillna(0)
                    rk["증감"] = rk["누적 판매량"] - rk["전년 동기"]
                    rk["증감률(%)"] = np.where(rk["전년 동기"] > 0, rk["증감"] / rk["전년 동기"] * 100, 0)
                    total = rk["누적 판매량"].sum()
                    rk["구성비(%)"] = (rk["누적 판매량"] / total * 100) if total > 0 else 0.0
                    rk["순위"] = rk["누적 판매량"].rank(method="min", ascending=False).astype(int)
                    rk["대상 업체수"] = len(rk)
                    rk["업종"] = cs_upjong_map(sub, rk.index, pd.Series(usage == "업무용", index=rk.index))
                    rk["구분"] = usage
                    rk = rk.rename_axis("고객명").reset_index()
                    return rk.sort_values(["순위", "고객명"]).reset_index(drop=True)

                rk_ind = cs_build_rank("산업용")
                rk_biz = cs_build_rank("업무용")

                # ── 검색한 고객의 순위 현황 ──
                if cs_found_names:
                    st.markdown(f"#### 🔎 검색 고객 순위 현황 ('{cs_query.strip()}' 검색 결과)")
                    parts = [r[r["고객명"].isin(cs_found_set)] for r in (rk_ind, rk_biz) if not r.empty]
                    found = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
                    if found.empty:
                        st.info("검색된 고객은 산업용·업무용 랭킹 대상이 아닙니다.")
                    else:
                        found = found.sort_values(["구분", "순위", "고객명"]).reset_index(drop=True)
                        found["Top30"] = np.where(found["순위"] <= 30, "🏅 Top30", "")
                        found_cols = ["구분", "순위", "대상 업체수", "고객명", "업종", "누적 판매량",
                                      "전년 동기", "증감", "증감률(%)", "구성비(%)", "Top30"]

                        def cs_style_found(row):
                            css = "background-color: #fff3bf; font-weight: 600;" if row["순위"] <= 30 else ""
                            return [css] * len(row)

                        cs_show_df(found[found_cols], cs_style_found, CS_FMT, key=f"cs_found{key_sfx}")
                        n_top = int((found["순위"] <= 30).sum())
                        st.caption(f"※ 검색 결과 중 Top30에 포함된 업체: {n_top}개 (노란 배경)")
                        n_out = len(cs_found_set - set(found["고객명"]))
                        if n_out > 0:
                            st.caption(f"※ 산업용·업무용 랭킹 대상이 아니어서 제외된 검색 고객: {n_out}개")
                    st.markdown("<br>", unsafe_allow_html=True)

                # ── 산업용 / 업무용 Top30 ──
                def cs_style_rank(row):
                    if row["고객명"] == cs_sel_name:
                        css = "background-color: #ffd8a8; font-weight: 700;"
                    elif row["고객명"] in cs_found_set:
                        css = "background-color: #ffe8cc;"
                    else:
                        css = ""
                    return [css] * len(row)

                rank_cols = ["순위", "고객명", "업종", "누적 판매량", "전년 동기", "증감률(%)"]
                cs_r1, cs_r2 = st.columns(2)
                for cs_col, cs_title, cs_rk, cs_k in ((cs_r1, "🏭 산업용 Top 30", rk_ind, "ind"),
                                                      (cs_r2, "🏢 업무용 Top 30", rk_biz, "biz")):
                    with cs_col:
                        if cs_rk.empty:
                            st.markdown(f"#### {cs_title}")
                            st.info("해당 용도의 데이터가 없습니다.")
                        else:
                            st.markdown(f"#### {cs_title} <span style='font-size:13px; font-weight:normal; color:gray;'>"
                                        f"(전체 {len(cs_rk)}개 업체)</span>", unsafe_allow_html=True)
                            cs_top = cs_rk.head(30)[rank_cols].reset_index(drop=True)
                            cs_show_df(cs_top, cs_style_rank, CS_FMT, key=f"cs_rank_{cs_k}{key_sfx}",
                                       height=min(35 * (len(cs_top) + 1) + 3, 1100))
            else:
                st.info("랭킹을 만들 업종별 CSV 데이터가 없습니다.")
            continue


        # ─────────────────────────────────────────────────────────
        # 1. 지도
        # ─────────────────────────────────────────────────────────
        st.markdown(
            f"### 🗺️ 1. 대용량 수요처 이상 감지 모니터링 지도 "
            f"<span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>",
            unsafe_allow_html=True
        )
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

        map_c1, map_c2, map_c3 = st.columns([1, 1, 1])
        with map_c1:
            map_usage = st.radio("📍 지도에 표시할 용도 선택", ["산업용", "업무용"], index=0, horizontal=True, key=f"map_radio_{key_sfx}")
        with map_c2:
            comp_mode = st.radio("📍 비교 기준", ["YoY", "전월대비"], index=0, horizontal=True, key=f"comp_mode_{key_sfx}")
        with map_c3:
            map_style_ui = st.radio("📍 지도 배경 테마", ["다크 모드 (기본)", "일반 도로 지도"], index=0, horizontal=True, key=f"map_style_{key_sfx}")

        # ✅ Carto 스타일 (Mapbox 토큰 불필요)
        deck_map_style = (
            "https://basemaps.cartocdn.com/gl/dark-matter-gl-style/style.json"
            if map_style_ui == "다크 모드 (기본)"
            else "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json"
        )

        curr_year = sel_year_rpt
        curr_month = max_month
        if comp_mode == "YoY":
            prev_year = curr_year - 1
            prev_month = curr_month
        else:
            prev_year = curr_year if curr_month > 1 else curr_year - 1
            prev_month = curr_month - 1 if curr_month > 1 else 12

        def get_mask(df, y, m, agg):
            if agg == "누적 실적 (1월~당월)":
                return (df["연_csv"] == y) & (df["월_csv"] <= m)
            else:
                return (df["연_csv"] == y) & (df["월_csv"] == m)

        if (not df_csv_tab.empty and "도로명주소" in df_csv_tab.columns
                and "고객명" in df_csv_tab.columns
                and val_col in df_csv_tab.columns
                and "용도" in df_csv_tab.columns):

            if map_usage == "산업용":
                df_map_filtered = df_csv_tab[df_csv_tab["용도"] == "산업용"].copy()
            else:
                if "상품명" in df_csv_tab.columns:
                    prod_s = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                    mask_u = (df_csv_tab["용도"] == "업무용") | (prod_s.isin(["냉난방용(업무)", "업무난방용", "주한미군"]))
                    df_map_filtered = df_csv_tab[mask_u].copy()
                else:
                    df_map_filtered = df_csv_tab[df_csv_tab["용도"] == "업무용"].copy()

            df_map_filtered["용도_태그"] = f"[{map_usage}]"
            mask_curr_map = get_mask(df_map_filtered, curr_year, curr_month, agg_mode)
            mask_prev_map = get_mask(df_map_filtered, prev_year, prev_month, agg_mode)

            map_curr = (df_map_filtered[mask_curr_map]
                        .groupby(["고객명", "도로명주소", "용도_태그"], as_index=False)[val_col]
                        .sum().rename(columns={val_col: "당해년도"}))
            map_prev = (df_map_filtered[mask_prev_map]
                        .groupby(["고객명", "도로명주소", "용도_태그"], as_index=False)[val_col]
                        .sum().rename(columns={val_col: "전년도"}))

            if not map_curr.empty and not map_prev.empty:
                df_map_merged = pd.merge(map_curr, map_prev, on=["고객명", "도로명주소", "용도_태그"], how="inner").fillna(0)
                df_map_merged["증감률(%)"] = np.where(
                    df_map_merged["전년도"] > 0,
                    ((df_map_merged["당해년도"] - df_map_merged["전년도"]) / df_map_merged["전년도"]) * 100,
                    0
                )
                alarm_df = df_map_merged[df_map_merged["증감률(%)"] <= -5].copy()

                if alarm_df.empty:
                    st.success(f"✅ 선택한 기간 내 5% 이상 하락한 {map_usage} 리스크 업체가 없습니다.")
                else:
                    st.warning(f"🚨 총 **{len(alarm_df)}**개의 {map_usage} 업체에서 5% 이상 하락 신호가 감지되었습니다.")

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
                            if rate <= -20:   level, c_color, r_size = "심각", [180, 0, 0, 255], 150
                            elif rate <= -10: level, c_color, r_size = "경계", [255, 80, 80, 200], 100
                            else:             level, c_color, r_size = "주의", [255, 150, 150, 200], 80
                        else:
                            if rate <= -20:   level, c_color, r_size = "심각", [0, 0, 180, 255], 150
                            elif rate <= -10: level, c_color, r_size = "경계", [80, 150, 255, 200], 100
                            else:             level, c_color, r_size = "주의", [120, 180, 255, 200], 80
                        colors.append(c_color)
                        radiuses.append(r_size)
                        info = f"<b>{row['용도_태그']} {row['고객명']} <span style='color:red;'>[{level}]</span></b><br/>"
                        info += f"전년/전월: {row['전년도']:,.0f} / 당해: {row['당해년도']:,.0f}<br/>"
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
                        editor_key = f"editor_{map_usage}_{key_sfx}"
                        selected_indices = []
                        if editor_key in st.session_state:
                            edited_rows = st.session_state[editor_key].get("edited_rows", {})
                            for idx_str, row_changes in edited_rows.items():
                                if row_changes.get("선택", False):
                                    idx_int = int(idx_str)
                                    if idx_int < len(alarm_df):
                                        selected_indices.append(idx_int)

                        map_df = alarm_df.iloc[selected_indices].copy() if selected_indices else alarm_df.copy()

                        layer = pdk.Layer(
                            "ScatterplotLayer", data=map_df,
                            get_position='[lon, lat]', get_color='color', get_radius='radius',
                            pickable=True, opacity=0.6, filled=True, stroked=True,
                            get_line_color=[255, 255, 255, 200], line_width_min_pixels=1, radius_max_pixels=40
                        )
                        layers = [layer]
                        start_lat, start_lon = 35.8660194, 128.5332943

                        if selected_indices:
                            # ✅ use_container_width=True (1.63.0에서 유효한 파라미터)
                            draw_route = st.button(
                                "🚗 선택 업체 최적 동선(실제 도로) 그리기",
                                use_container_width=True,
                                key=f"draw_route_btn_{key_sfx}"
                            )
                            start_pt_data = pd.DataFrame([{
                                "lon": start_lon, "lat": start_lat,
                                "tooltip": "<b>🏢 대성에너지 서부지사 (출발지)</b>",
                                "color": [255, 193, 7, 255], "radius": 150
                            }])
                            layers.append(pdk.Layer(
                                "ScatterplotLayer", data=start_pt_data,
                                get_position='[lon, lat]', get_color='color', get_radius='radius',
                                pickable=True, opacity=1.0, filled=True, stroked=True,
                                get_line_color=[0, 0, 0, 255], line_width_min_pixels=2, radius_max_pixels=15
                            ))
                            if draw_route:
                                with st.spinner("카카오 모빌리티 API를 통해 최적 도로 경로를 탐색 중입니다..."):
                                    unvisited = map_df[['lon', 'lat', '고객명']].to_dict('records')
                                    curr_lat_r, curr_lon_r = start_lat, start_lon
                                    ordered_stops = []
                                    while unvisited:
                                        nearest = min(unvisited, key=lambda pt: (pt['lat'] - curr_lat_r) ** 2 + (pt['lon'] - curr_lon_r) ** 2)
                                        ordered_stops.append(nearest)
                                        curr_lat_r, curr_lon_r = nearest['lat'], nearest['lon']
                                        unvisited.remove(nearest)
                                    full_route_coords = []
                                    current_pt = [start_lon, start_lat]
                                    for stop in ordered_stops:
                                        target_pt = [stop['lon'], stop['lat']]
                                        route_segment = get_kakao_route(current_pt[0], current_pt[1], target_pt[0], target_pt[1])
                                        full_route_coords.extend(route_segment if route_segment else [current_pt, target_pt])
                                        current_pt = target_pt
                                    if full_route_coords:
                                        path_data = pd.DataFrame([{"path": full_route_coords, "color": [46, 204, 113, 255]}])
                                        layers.append(pdk.Layer(
                                            "PathLayer", data=path_data,
                                            get_path="path", get_color="color",
                                            width_scale=20, width_min_pixels=3, get_width=5
                                        ))
                                        st.success("✨ 최적 도로 경로가 지도에 표시되었습니다!")

                        view_state = pdk.ViewState(
                            latitude=start_lat if selected_indices else alarm_df['lat'].mean(),
                            longitude=start_lon if selected_indices else alarm_df['lon'].mean(),
                            zoom=11, pitch=40,
                        )
                        deck_obj = pdk.Deck(
                            map_style=deck_map_style,
                            layers=layers,
                            initial_view_state=view_state,
                            tooltip={"html": "{tooltip}", "style": {"backgroundColor": "white", "color": "black"}}
                        )
                        st.pydeck_chart(deck_obj)

                        st.markdown(
                            f"<br><b>📋 지도 표기 업체 요약표</b> "
                            f"<span style='font-size:13px; color:#d32f2f; margin-left:10px;'>✅ 표 좌측 [선택] 체크 시 상단 지도에 선택 업체와 출발지가 뜹니다.</span>",
                            unsafe_allow_html=True
                        )
                        prev_col_name = "전년도" if comp_mode == "YoY" else "전월"
                        curr_col_name = "당해년도" if comp_mode == "YoY" else "당월"
                        show_cols = ['용도_태그', '고객명', '도로명주소', '전년도', '당해년도', '증감', '증감률(%)']
                        df_show = alarm_df[show_cols].copy()
                        df_show = df_show.rename(columns={"전년도": prev_col_name, "당해년도": curr_col_name})
                        df_show.insert(0, "No.", range(1, len(df_show) + 1))
                        df_show.insert(0, "선택", False)
                        df_show["비고"] = np.where(df_show["증감률(%)"] <= -99.9, "폐업의심", "")

                        sum_prev_all = df_show[prev_col_name].sum()
                        sum_curr_all = df_show[curr_col_name].sum()
                        sum_rate_all = ((sum_curr_all - sum_prev_all) / sum_prev_all * 100) if sum_prev_all > 0 else 0
                        total_row = pd.DataFrame([{
                            "선택": False, "No.": "", "용도_태그": "💡 총계",
                            "고객명": "", "도로명주소": "",
                            prev_col_name: sum_prev_all, curr_col_name: sum_curr_all,
                            "증감": sum_curr_all - sum_prev_all,
                            "증감률(%)": sum_rate_all, "비고": ""
                        }])
                        df_show = pd.concat([df_show, total_row], ignore_index=True)

                        # ✅ st.data_editor에 순수 DataFrame만 전달 (Styler 절대 사용 안 함)
                        st.data_editor(
                            df_show,
                            column_config={
                                "선택": st.column_config.CheckboxColumn("선택", default=False),
                                prev_col_name: st.column_config.NumberColumn(format="%.0f"),
                                curr_col_name: st.column_config.NumberColumn(format="%.0f"),
                                "증감": st.column_config.NumberColumn(format="%.0f"),
                                "증감률(%)": st.column_config.NumberColumn(format="%.1f"),
                            },
                            disabled=[c for c in df_show.columns if c != "선택"],
                            hide_index=True,
                            use_container_width=True,
                            key=editor_key,
                        )
                    else:
                        st.error("매핑된 위경도 좌표가 없어 지도를 표시할 수 없습니다.")
            else:
                st.info("비교할 과거 또는 당해 연도 데이터가 없습니다.")
        else:
            st.info("데이터에 '도로명주소', '고객명', '용도' 컬럼이 없거나 데이터가 부족합니다.")

        st.markdown("<hr style='border-top: 2px solid #1e3a8a; margin: 50px 0 20px 0;'>", unsafe_allow_html=True)

        # ─────────────────────────────────────────────────────────
        # 분석 함수
        # ─────────────────────────────────────────────────────────
        def render_full_usage_report(usage_name, section_num, key_sfx_inner, db_key):
            df_u_csv = pd.DataFrame()

            st.markdown(f"#### 📈 {section_num}. 용도별 판매량 분석 : {usage_name}")

            if not df_long_rpt.empty:
                df_u = df_long_rpt[(df_long_rpt["그룹"] == usage_name)]
                p_curr_act_all = df_u[(df_u["연"] == curr_year) & (df_u["계획/실적"] == "실적")].groupby("월")["값"].sum()
                p_prev_act_all = df_u[(df_u["연"] == prev_year) & (df_u["계획/실적"] == "실적")].groupby("월")["값"].sum()
            else:
                if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                    if "상품명" in df_csv_tab.columns:
                        csv_products = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                    else:
                        csv_products = pd.Series([""] * len(df_csv_tab), index=df_csv_tab.index)
                    if usage_name == "산업용":
                        df_u_csv = df_csv_tab[csv_products == "산업용"].copy()
                    else:
                        valid_biz = ["냉난방용(업무)", "업무난방용", "주한미군"]
                        df_u_csv = df_csv_tab[csv_products.isin(valid_biz)].copy()
                    p_curr_act_all = df_u_csv[df_u_csv["연_csv"] == curr_year].groupby("월_csv")[val_col].sum()
                    p_prev_act_all = df_u_csv[df_u_csv["연_csv"] == prev_year].groupby("월_csv")[val_col].sum()
                else:
                    p_curr_act_all = pd.Series(dtype=float)
                    p_prev_act_all = pd.Series(dtype=float)

            months_list = list(range(1, curr_month + 1))

            if comp_mode == "YoY":
                if agg_mode == "누적 실적 (1월~당월)":
                    sum_act = p_curr_act_all[p_curr_act_all.index <= curr_month].sum()
                    sum_prev = p_prev_act_all[p_prev_act_all.index <= prev_month].sum()
                    top_title = f"**■ 누적 실적 비교 ({curr_month}월 누적)**"
                else:
                    sum_act = p_curr_act_all.get(curr_month, 0)
                    sum_prev = p_prev_act_all.get(prev_month, 0)
                    top_title = f"**■ 당월 실적 비교 ({curr_month}월 당월)**"
                prev_name = f"{prev_year}년"
                curr_name = f"{curr_year}년"
                diff_label = "전년대비"
                vals_prev = [p_prev_act_all.get(m, 0) for m in months_list]
                prev_legend = f"{prev_year}년 실적"
            else:
                if agg_mode == "누적 실적 (1월~당월)":
                    sum_act = p_curr_act_all[p_curr_act_all.index <= curr_month].sum()
                    sum_prev = p_prev_act_all[p_prev_act_all.index <= prev_month].sum()
                    top_title = f"**■ 누적 ({prev_month}월 vs {curr_month}월)**"
                else:
                    sum_act = p_curr_act_all.get(curr_month, 0)
                    sum_prev = p_prev_act_all.get(prev_month, 0)
                    top_title = f"**■ 전월 실적 비교 ({prev_month}월 vs {curr_month}월)**"
                prev_name = f"전월({prev_month}월)"
                curr_name = f"당월({curr_month}월)"
                diff_label = "전월대비"
                vals_prev = []
                for m in months_list:
                    if m > 1:
                        vals_prev.append(p_curr_act_all.get(m - 1, 0))
                    else:
                        if not df_long_rpt.empty:
                            p_last_yr = df_long_rpt[
                                (df_long_rpt["그룹"] == usage_name) &
                                (df_long_rpt["연"] == curr_year - 1) &
                                (df_long_rpt["계획/실적"] == "실적")
                            ].groupby("월")["값"].sum()
                            vals_prev.append(p_last_yr.get(12, 0))
                        elif not df_u_csv.empty:
                            p_last_yr = df_u_csv[df_u_csv["연_csv"] == curr_year - 1].groupby("월_csv")[val_col].sum()
                            vals_prev.append(p_last_yr.get(12, 0))
                        else:
                            vals_prev.append(0)
                prev_legend = "전월 실적"

            diff_prev = sum_act - sum_prev
            rate_prev = (sum_act / sum_prev * 100) if sum_prev > 0 else 0
            sign_prev = "+" if diff_prev > 0 else ""
            curr_legend = f"{curr_year}년 실적" if comp_mode == "YoY" else "당월 실적"
            desc_status = "감소" if diff_prev < 0 else "증가"

            st.markdown(
                f"""<div style="background-color: #f8f9fa; border-left: 5px solid #1e3a8a; padding: 15px;
                    margin-bottom: 20px; border-radius: 4px;">
                    <div style="font-size: 15px; color: #1e3a8a; font-weight: 700; line-height: 1.6;">
                        💡 [요약] 당해 실적: {sum_act:,.0f} {unit_str}<br>
                        {diff_label} <span style="color: {'#d32f2f' if diff_prev < 0 else '#1f77b4'};">
                        {abs(diff_prev):,.0f} {unit_str} {desc_status} ({sign_prev}{rate_prev:.1f}%)</span>
                    </div></div>""",
                unsafe_allow_html=True
            )

            vals_act = [p_curr_act_all.get(m, 0) for m in months_list]
            overall_max = max(max(vals_act) if vals_act else 0, max(vals_prev) if vals_prev else 0,
                              sum_prev, sum_act)
            yaxis_range = [0, overall_max * 1.1 if overall_max > 0 else 100]
            # 좌측(누적/당월 비교)·우측(월별 추이) 차트가 각자 자기 값에 맞는 눈금을 쓰도록 분리
            _max_c = max(sum_prev, sum_act)
            _max_m = max(max(vals_act) if vals_act else 0, max(vals_prev) if vals_prev else 0)
            yaxis_range_c = [0, _max_c * 1.1 if _max_c > 0 else 100]
            yaxis_range_m = [0, _max_m * 1.1 if _max_m > 0 else 100]

            col_c, col_m = st.columns([1, 2.5])
            with col_c:
                st.markdown(top_title, unsafe_allow_html=True)
                fig_c = go.Figure()
                fig_c.update_layout(margin=dict(t=30, b=20, l=40, r=10), height=420, showlegend=False)
                fig_c.update_yaxes(range=yaxis_range_c)
                fig_c.add_trace(go.Bar(
                    x=[f"{prev_name}<br>실적", f"{curr_name}<br>실적"], y=[sum_prev, sum_act],
                    marker_color=[COLOR_PREV, COLOR_ACT],
                    text=[f"{sum_prev:,.0f}", f"{sum_act:,.0f}"],
                    textposition='auto', textfont=dict(size=14)
                ))
                # ✅ use_container_width=True (1.63.0 정상 파라미터)
                st.plotly_chart(fig_c, key=f"fig_c_{usage_name}_{key_sfx_inner}", use_container_width=True)

            with col_m:
                st.markdown(f"**■ 월별 실적 추이**", unsafe_allow_html=True)
                fig_m = go.Figure()
                fig_m.update_layout(
                    barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1),
                    margin=dict(t=30, b=20, l=40, r=10), height=420,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                fig_m.update_yaxes(range=yaxis_range_m)
                fig_m.add_trace(go.Bar(x=months_list, y=vals_prev, name=prev_legend, marker_color=COLOR_PREV,
                    text=[f"{v:,.0f}" if v > 0 else "" for v in vals_prev], textposition='auto', textfont=dict(size=11)))
                fig_m.add_trace(go.Bar(x=months_list, y=vals_act, name=curr_legend, marker_color=COLOR_ACT,
                    text=[f"{v:,.0f}" if v > 0 else "" for v in vals_act], textposition='auto', textfont=dict(size=11)))
                # ✅ use_container_width=True
                st.plotly_chart(fig_m, key=f"fig_m_{usage_name}_{key_sfx_inner}", use_container_width=True)

            if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                if "상품명" in df_csv_tab.columns:
                    csv_products2 = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                else:
                    csv_products2 = pd.Series([""] * len(df_csv_tab), index=df_csv_tab.index)

                if usage_name == "산업용":
                    df_sub = df_csv_tab[csv_products2 == "산업용"].copy()
                    grp_col = "업종"
                else:
                    valid_biz = ["냉난방용(업무)", "업무난방용", "주한미군"]
                    df_sub = df_csv_tab[csv_products2.isin(valid_biz)].copy()
                    if "업종분류" in df_sub.columns:
                        df_sub["업종"] = df_sub["업종분류"]
                    grp_col = "업종"

                mask_curr_sub = get_mask(df_sub, curr_year, curr_month, agg_mode)
                mask_prev_sub = get_mask(df_sub, prev_year, prev_month, agg_mode)

                if not df_sub.empty and grp_col in df_sub.columns:
                    curr_ind = (df_sub[mask_curr_sub].groupby(grp_col, as_index=False)[val_col]
                                .sum().rename(columns={val_col: curr_name}))
                    prev_ind = (df_sub[mask_prev_sub].groupby(grp_col, as_index=False)[val_col]
                                .sum().rename(columns={val_col: prev_name}))
                    ind_comp = pd.merge(prev_ind, curr_ind, on=grp_col, how="outer").fillna(0)
                    ind_comp = ind_comp.sort_values(curr_name, ascending=False).reset_index(drop=True)

                    if len(ind_comp) > 10:
                        top10 = ind_comp.iloc[:10].copy()
                        others = ind_comp.iloc[10:].copy()
                        others_row = pd.DataFrame([{grp_col: "기타", prev_name: others[prev_name].sum(), curr_name: others[curr_name].sum()}])
                        ind_comp_plot = pd.concat([top10, others_row], ignore_index=True)
                    else:
                        ind_comp_plot = ind_comp.copy()

                    ind_comp_plot["증감절대값"] = abs(ind_comp_plot[curr_name] - ind_comp_plot[prev_name])
                    max_diff_idx = ind_comp_plot["증감절대값"].idxmax()
                    colors_act = [COLOR_ACT] * len(ind_comp_plot)
                    if pd.notna(max_diff_idx): colors_act[int(max_diff_idx)] = "#d32f2f"

                    st.markdown(f"**■ 세부 업종별 판매량 비교**")
                    fig_ind = go.Figure()
                    fig_ind.add_trace(go.Bar(x=ind_comp_plot[grp_col], y=ind_comp_plot[prev_name], name=prev_name,
                        marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v > 0 else "" for v in ind_comp_plot[prev_name]],
                        textposition='auto', textfont=dict(size=11)))
                    fig_ind.add_trace(go.Bar(x=ind_comp_plot[grp_col], y=ind_comp_plot[curr_name], name=curr_name,
                        marker_color=colors_act, text=[f"{v:,.0f}" if v > 0 else "" for v in ind_comp_plot[curr_name]],
                        textposition='auto', textfont=dict(size=11)))
                    fig_ind.update_layout(barmode='group', margin=dict(t=10, b=10, l=10, r=10), height=420,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                    # ✅ use_container_width=True
                    st.plotly_chart(fig_ind, key=f"fig_ind_{usage_name}_{key_sfx_inner}", use_container_width=True)

                st.markdown("<hr style='border-top: 1px dashed #ccc; margin: 30px 0;'>", unsafe_allow_html=True)
                st.markdown(f"**🔍 {usage_name} 개별 고객 상세 차트**")

                if not df_sub.empty and "고객명" in df_sub.columns and grp_col in df_sub.columns:
                    c_curr_all = (df_sub[mask_curr_sub].groupby(["고객명", grp_col], as_index=False)[val_col]
                                  .sum().rename(columns={val_col: curr_name}))
                    c_prev_all = (df_sub[mask_prev_sub].groupby(["고객명", grp_col], as_index=False)[val_col]
                                  .sum().rename(columns={val_col: prev_name}))
                    grp_top = pd.merge(c_prev_all, c_curr_all, on=["고객명", grp_col], how="outer").fillna(0)
                    grp_top = grp_top.sort_values(curr_name, ascending=False).reset_index(drop=True)
                    grp_top = grp_top[(grp_top[curr_name] > 0) | (grp_top[prev_name] > 0)].reset_index(drop=True)

                    top_customers = [c for c in grp_top["고객명"] if "💡" not in str(c)]
                    sel_cust = st.selectbox(
                        f"상세 분석할 고객명을 선택하세요 ({usage_name})",
                        ["선택 안함"] + top_customers,
                        key=f"sel_cust_{usage_name}_{key_sfx_inner}"
                    )

                    if sel_cust != "선택 안함":
                        c_data = df_csv_tab[df_csv_tab["고객명"] == sel_cust]
                        mask_c_curr = get_mask(c_data, curr_year, curr_month, agg_mode)
                        mask_c_prev = get_mask(c_data, prev_year, prev_month, agg_mode)
                        sum_cur_c = c_data[mask_c_curr][val_col].sum()
                        sum_prev_c = c_data[mask_c_prev][val_col].sum()
                        chart_title = (f"'{sel_cust}' 누적 사용량 ({curr_month}월 누적)"
                                       if agg_mode == "누적 실적 (1월~당월)"
                                       else f"'{sel_cust}' 당월 사용량 ({curr_month}월)")
                        diff_val = sum_cur_c - sum_prev_c
                        rate_val = (sum_cur_c / sum_prev_c * 100) if sum_prev_c > 0 else 0
                        sign_str = "+" if diff_val > 0 else ""
                        yoy_text = f"{diff_label} 증감: {sign_str}{diff_val:,.0f} ({rate_val:.1f}%)"

                        cc1, cc2 = st.columns([1, 2])
                        with cc1:
                            fig_cust_cum = go.Figure()
                            fig_cust_cum.update_layout(title=chart_title, margin=dict(t=50, b=20, l=40, r=10), height=350)
                            fig_cust_cum.add_trace(go.Bar(
                                x=[prev_name, curr_name], y=[sum_prev_c, sum_cur_c],
                                marker_color=[COLOR_PREV, COLOR_ACT],
                                text=[f"{sum_prev_c:,.0f}", f"{sum_cur_c:,.0f}"],
                                textposition='auto'
                            ))
                            fig_cust_cum.add_annotation(
                                x=0.5, y=1.05, xref="paper", yref="paper",
                                text=f"<b>{yoy_text}</b>", showarrow=False,
                                font=dict(size=13, color="#d32f2f" if diff_val < 0 else "#1f77b4"),
                                bgcolor="#f8f9fa", bordercolor="#d0d7e5", borderwidth=1, borderpad=4
                            )
                            # ✅ use_container_width=True
                            st.plotly_chart(fig_cust_cum, key=f"fig_cum_{usage_name}_{key_sfx_inner}_{sel_cust}", use_container_width=True)

                        with cc2:
                            fig_cust_mon = go.Figure()
                            fig_cust_mon.update_layout(
                                title=f"'{sel_cust}' 월별 사용량 추이", barmode='group',
                                xaxis=dict(tickmode='linear', tick0=1, dtick=1),
                                margin=dict(t=50, b=20, l=40, r=10), height=350,
                                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                            )
                            cur_vals_c = [c_data[(c_data['연_csv'] == curr_year) & (c_data['월_csv'] == m)][val_col].sum() for m in months_list]
                            if comp_mode == "YoY":
                                prev_vals_c = [c_data[(c_data['연_csv'] == prev_year) & (c_data['월_csv'] == m)][val_col].sum() for m in months_list]
                            else:
                                prev_vals_c = []
                                for m in months_list:
                                    if m > 1:
                                        prev_vals_c.append(c_data[(c_data['연_csv'] == curr_year) & (c_data['월_csv'] == m - 1)][val_col].sum())
                                    else:
                                        prev_vals_c.append(c_data[(c_data['연_csv'] == curr_year - 1) & (c_data['월_csv'] == 12)][val_col].sum())
                            fig_cust_mon.add_trace(go.Bar(x=months_list, y=prev_vals_c, name=prev_legend,
                                marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v > 0 else "" for v in prev_vals_c],
                                textposition='auto', textfont=dict(size=11)))
                            fig_cust_mon.add_trace(go.Bar(x=months_list, y=cur_vals_c, name=curr_legend,
                                marker_color=COLOR_ACT, text=[f"{v:,.0f}" if v > 0 else "" for v in cur_vals_c],
                                textposition='auto', textfont=dict(size=11)))
                            # ✅ use_container_width=True
                            st.plotly_chart(fig_cust_mon, key=f"fig_mon_{usage_name}_{key_sfx_inner}_{sel_cust}", use_container_width=True)

        render_full_usage_report("산업용", "2", key_sfx, "ind")
        st.markdown("<hr style='margin: 50px 0; border-top: 2px solid #ccc;'>", unsafe_allow_html=True)
        render_full_usage_report("업무용", "3", key_sfx, "biz")

        # ─────────────────────────────────────────────────────────
        # 보고서 출력
        # ─────────────────────────────────────────────────────────
        st.markdown("<hr style='border-top: 2px solid #bbb; margin: 40px 0 20px 0;'>", unsafe_allow_html=True)
        st.markdown("### 🖨️ 4. 보고서 출력")
        st.markdown("""
            <style>
            @media print {
                header[data-testid="stHeader"] { display: none !important; }
                section[data-testid="stSidebar"] { display: none !important; }
                div[data-testid="stToolbar"] { display: none !important; }
            }
            </style>
        """, unsafe_allow_html=True)
        st.html("""
            <button onclick="window.parent.print()"
                style="padding: 12px 20px; font-size: 16px; border-radius: 8px;
                       background-color: #1e3a8a; color: white; border: none; cursor: pointer;
                       width: 100%; font-weight: bold; margin: 2px;">
                🖨️ 현재 화면 전체를 PDF로 다운로드 (인쇄)
            </button>
        """)
