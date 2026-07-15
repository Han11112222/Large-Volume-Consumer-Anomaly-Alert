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
# 코멘트 DB 저장
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
        dict(selector="th", props=[("text-align", "center"), ("vertical-align", "middle"),
                                   ("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")]),
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
                            path_coords.append([vertices[i], vertices[i+1]])
            return path_coords
        else:
            return None
    except Exception:
        return None


# ─────────────────────────────────────────────────────────
# ✅ [핵심 수정] CSV 날짜 파싱 - @st.cache_data 제거, 일반 함수로
# ─────────────────────────────────────────────────────────
def preprocess_csv(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    원본 CSV를 받아서 날짜 파싱(연_csv, 월_csv)만 추가.
    사용량 단위 변환은 절대 여기서 하지 않음.
    """
    if df_raw.empty:
        return df_raw.copy()

    df = df_raw.copy()

    # 날짜 파싱
    date_col = None
    for c in ["청구년월", "매출년월", "년월", "기준년월"]:
        if c in df.columns:
            date_col = c
            break

    parsed = pd.Series([pd.NaT] * len(df), index=df.index)

    if date_col:
        try:
            p1 = pd.to_datetime(df[date_col], format="%b-%y", errors="coerce")
            p2 = pd.to_datetime(df[date_col], format="%Y%m", errors="coerce")
            p3 = pd.to_datetime(df[date_col], errors="coerce")
            # 순서대로 우선 적용
            parsed = p1.where(p1.notna(), p2).where(p1.notna() | p2.notna(), p3)
        except Exception:
            pass

    fallback = pd.to_datetime("2026-03-01")
    df["날짜_파싱"] = parsed.fillna(fallback)
    df["연_csv"] = df["날짜_파싱"].dt.year.astype(int)
    df["월_csv"] = df["날짜_파싱"].dt.month.astype(int)

    return df


def get_unit_series(df: pd.DataFrame, unit_str: str) -> pd.Series:
    """
    단위 변환된 사용량 Series를 반환. df 원본은 절대 수정하지 않음.
    """
    if unit_str == "GJ" and "사용량(mj)" in df.columns:
        return pd.to_numeric(df["사용량(mj)"], errors="coerce").fillna(0.0) / 1000.0
    elif unit_str == "천m³" and "사용량(m3)" in df.columns:
        return pd.to_numeric(df["사용량(m3)"], errors="coerce").fillna(0.0) / 1000.0
    # fallback
    for c in ["사용량(mj)", "사용량(m3)"]:
        if c in df.columns:
            return pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return pd.Series([0.0] * len(df), index=df.index)


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
        up_csvs = st.file_uploader("가정용외_*.csv 형식 (다중 업로드 가능)", type=["csv"],
                                   accept_multiple_files=True, key="csv_uploader")
        if up_csvs:
            df_list = []
            for f in up_csvs:
                df = load_safe_csv(f.getvalue())
                if not df.empty:
                    df_list.append(df)
            if df_list:
                st.session_state['merged_csv_df'] = pd.concat(df_list, ignore_index=True)
        else:
            if 'merged_csv_df' in st.session_state:
                del st.session_state['merged_csv_df']

    # GitHub API 로드 결과 표시
    if 'github_csv_loaded' in st.session_state:
        names = st.session_state['github_csv_loaded']
        st.success(f"✅ GitHub CSV {len(names)}개 로드 완료")
        with st.expander("로드된 파일 목록 보기"):
            for n in names:
                st.write(f"- {n}")
    if 'github_csv_error' in st.session_state:
        st.error(f"GitHub 로드 오류: {st.session_state['github_csv_error']}")

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

# ✅ CSV 원본 로드 (단위 변환 없이, 숫자 정제만)
df_csv_raw = pd.DataFrame()

def load_csvs_from_github() -> pd.DataFrame:
    """
    GitHub API를 통해 레포의 가정용외_*.csv 파일을 모두 읽어 합칩니다.
    glob 방식보다 안정적으로 모든 파일을 가져옵니다.
    """
    try:
        token = ""
        try:
            token = st.secrets.get("GITHUB_TOKEN", "")
        except Exception:
            pass

        if not token:
            return pd.DataFrame()

        g = Github(token)
        repo = g.get_repo(REPO_NAME)
        contents = repo.get_contents("")  # 루트 디렉토리 파일 목록

        csv_files = [
            c for c in contents
            if c.name.endswith(".csv") and "가정용외" in c.name
        ]
        csv_files = sorted(csv_files, key=lambda c: c.name)

        loaded_names = []
        csv_list = []
        for file_content in csv_files:
            try:
                raw_bytes = file_content.decoded_content
                for enc in ["utf-8-sig", "cp949", "utf-8"]:
                    try:
                        temp_df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, thousands=',')
                        temp_df.columns = temp_df.columns.str.strip()
                        csv_list.append(temp_df)
                        loaded_names.append(file_content.name)
                        break
                    except Exception:
                        pass
            except Exception:
                pass

        if csv_list:
            # session_state에 로드 결과 저장 (사이드바에서 표시용)
            st.session_state['github_csv_loaded'] = loaded_names
            return pd.concat(csv_list, ignore_index=True)
        else:
            return pd.DataFrame()

    except Exception as e:
        st.session_state['github_csv_error'] = str(e)
        return pd.DataFrame()

if src_csv == "레포 파일 사용":
    # ① 먼저 로컬 경로(glob) 시도
    repo_dir = Path(__file__).parent
    all_csvs = list(set(
        list(repo_dir.glob("*가정용외*.csv")) +
        list(repo_dir.glob("가정용외*.csv"))
    ))
    csv_list = []
    for p in sorted(all_csvs):
        for enc in ["utf-8-sig", "cp949", "utf-8"]:
            try:
                temp_df = pd.read_csv(p, encoding=enc, thousands=',')
                temp_df.columns = temp_df.columns.str.strip()
                csv_list.append(temp_df)
                break
            except Exception:
                pass
    if csv_list:
        df_csv_raw = pd.concat(csv_list, ignore_index=True)

    # ② 로컬에서 못 찾으면 GitHub API로 로드
    if df_csv_raw.empty:
        df_csv_raw = load_csvs_from_github()

if df_csv_raw.empty and 'merged_csv_df' in st.session_state:
    df_csv_raw = st.session_state['merged_csv_df'].copy()

# 숫자 정제 (원본 단위 그대로 보존)
if not df_csv_raw.empty:
    if "사용량(mj)" in df_csv_raw.columns:
        df_csv_raw["사용량(mj)"] = df_csv_raw["사용량(mj)"].apply(clean_korean_finance_number)
    if "사용량(m3)" in df_csv_raw.columns:
        df_csv_raw["사용량(m3)"] = df_csv_raw["사용량(m3)"].apply(clean_korean_finance_number)

# ✅ 날짜 파싱을 탭 루프 밖에서 한 번만 수행
if not df_csv_raw.empty:
    df_csv_parsed = preprocess_csv(df_csv_raw)
else:
    df_csv_parsed = pd.DataFrame()

comments_db = load_comments_db()

# ─────────────────────────────────────────────────────────
# 탭 루프
# ─────────────────────────────────────────────────────────
UNIT_VAL_COL = "사용량_단위변환"   # 탭별 단위변환 결과를 담을 임시 컬럼명

rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

for idx, rpt_tab in enumerate(rpt_tabs):
    with rpt_tab:
        if idx == 0:
            df_long_rpt = long_dict_rpt.get("열량", pd.DataFrame())
            unit_str = "GJ"
            key_sfx = "_gj"
        else:
            df_long_rpt = long_dict_rpt.get("부피", pd.DataFrame())
            unit_str = "천m³"
            key_sfx = "_vol"

        # ✅ [핵심] 탭별로 df_csv_parsed 복사 후, 단위변환 컬럼만 새로 추가
        #    원본 사용량 컬럼(mj/m3)은 절대 건드리지 않음
        if not df_csv_parsed.empty:
            df_csv_tab = df_csv_parsed.copy()
            df_csv_tab[UNIT_VAL_COL] = get_unit_series(df_csv_tab, unit_str)
        else:
            df_csv_tab = pd.DataFrame()

        val_col = UNIT_VAL_COL  # 이후 모든 집계는 이 컬럼 사용

        # ── 기준 일자 설정 ──
        st.markdown(f"#### 📅 기준 일자 설정")

        years_available = [2024, 2025, 2026]
        default_y_index = len(years_available) - 1
        default_m_index = 2

        if not df_long_rpt.empty:
            years_available = sorted(df_long_rpt["연"].unique().tolist())
            actual_data = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["값"] > 0)]
            if not actual_data.empty:
                max_year = int(actual_data["연"].max())
                max_month_val = int(actual_data[actual_data["연"] == max_year]["월"].max())
                default_y_index = years_available.index(max_year) if max_year in years_available else len(years_available) - 1
                default_m_index = max_month_val - 1

        if not df_csv_tab.empty and "날짜_파싱" in df_csv_tab.columns:
            csv_max_date = df_csv_tab["날짜_파싱"].max()
            if pd.notna(csv_max_date):
                y = int(csv_max_date.year)
                if y not in years_available:
                    years_available = sorted(set(years_available) | {y})
                default_y_index = years_available.index(y)
                default_m_index = int(csv_max_date.month) - 1

        c_y, c_m, c_agg, _ = st.columns([1, 1, 2, 1])
        with c_y:
            sel_year_rpt = st.selectbox("기준 연도", years_available, index=default_y_index, key=f"rpt_yr{key_sfx}")
        with c_m:
            sel_month_str = st.selectbox("기준 월", [f"{m}월" for m in range(1, 13)], index=default_m_index, key=f"rpt_mo{key_sfx}")
        with c_agg:
            agg_mode = st.radio("집계 기준", ["당월 실적", "누적 실적 (1월~당월)"], index=0, horizontal=True, key=f"agg_mode_{key_sfx}")

        max_month = int(sel_month_str.replace("월", ""))
        report_db_key = f"{sel_year_rpt}_{max_month}M_{unit_str}_yoy_only"

        if report_db_key not in comments_db:
            comments_db[report_db_key] = {}

        st.markdown("<hr style='margin: 10px 0 30px 0;'>", unsafe_allow_html=True)

        # ─────────────────────────────────────────────────────────
        # 1. 이상 감지 업체 지도 모니터링
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

        deck_map_style = "dark" if map_style_ui == "다크 모드 (기본)" else "road"

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

        if (not df_csv_tab.empty
                and "도로명주소" in df_csv_tab.columns
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
                    st.success(f"✅ 선택한 기간 내 YoY 5% 이상 하락한 {map_usage} 리스크 업체가 없습니다.")
                else:
                    st.warning(f"🚨 총 **{len(alarm_df)}**개의 {map_usage} 업체에서 5% 이상 하락 신호가 감지되었습니다. "
                               f"(지도에는 하락폭이 큰 주요 100개 업체를 표시합니다.)")

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
                            if rate <= -20:   level, c, r = "심각", [180, 0, 0, 255], 150
                            elif rate <= -10: level, c, r = "경계", [255, 80, 80, 200], 100
                            else:             level, c, r = "주의", [255, 150, 150, 200], 80
                        else:
                            if rate <= -20:   level, c, r = "심각", [0, 0, 180, 255], 150
                            elif rate <= -10: level, c, r = "경계", [80, 150, 255, 200], 100
                            else:             level, c, r = "주의", [120, 180, 255, 200], 80
                        colors.append(c); radiuses.append(r)
                        info = (f"<b>{row['용도_태그']} {row['고객명']} <span style='color:red;'>[{level}]</span></b><br/>"
                                f"전년/전월: {row['전년도']:,.0f} / 당해: {row['당해년도']:,.0f}<br/>"
                                f"증감률: <span style='color:red; font-weight:bold;'>{row['증감률(%)']:.1f}%</span><br/>"
                                f"<span style='font-size:0.8em; color:gray;'>{row['도로명주소']}</span>")
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
                            draw_route = st.button("🚗 선택 업체 최적 동선(실제 도로) 그리기",
                                                   width="stretch", key=f"draw_route_btn_{key_sfx}")
                            start_pt_data = pd.DataFrame([{
                                "lon": start_lon, "lat": start_lat,
                                "tooltip": "<b>🏢 대성에너지 서부지사 (출발지)</b><br>대구광역시 서구 와룡로73길 30",
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
                                    c_lat, c_lon = start_lat, start_lon
                                    ordered_stops = []
                                    while unvisited:
                                        nearest = min(unvisited, key=lambda pt: (pt['lat']-c_lat)**2 + (pt['lon']-c_lon)**2)
                                        ordered_stops.append(nearest)
                                        c_lat, c_lon = nearest['lat'], nearest['lon']
                                        unvisited.remove(nearest)
                                    full_route_coords = []
                                    current_pt = [start_lon, start_lat]
                                    for stop in ordered_stops:
                                        target_pt = [stop['lon'], stop['lat']]
                                        seg = get_kakao_route(current_pt[0], current_pt[1], target_pt[0], target_pt[1])
                                        full_route_coords.extend(seg if seg else [current_pt, target_pt])
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
                        r = pdk.Deck(
                            map_style=deck_map_style, layers=layers, initial_view_state=view_state,
                            tooltip={"html": "{tooltip}", "style": {"backgroundColor": "white", "color": "black", "font-family": "NanumGothic"}}
                        )
                        st.pydeck_chart(r)

                        st.markdown(
                            f"<br><b>📋 지도 표기 업체 요약표</b> "
                            f"<span style='font-size:13px; color:#d32f2f; margin-left:10px;'>✅ 표 좌측 [선택] 체크 시 상단 지도에 선택 업체와 출발지가 뜹니다.</span> "
                            f"<span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>",
                            unsafe_allow_html=True
                        )

                        prev_col_name = "전년도" if comp_mode == "YoY" else "전월"
                        curr_col_name = "당해년도" if comp_mode == "YoY" else "당월"

                        df_show = alarm_df[['용도_태그', '고객명', '도로명주소', '전년도', '당해년도', '증감', '증감률(%)']].copy()
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
                            "증감": sum_curr_all - sum_prev_all, "증감률(%)": sum_rate_all, "비고": ""
                        }])
                        df_show = pd.concat([df_show, total_row], ignore_index=True)

                        def highlight_map_total(s):
                            if s.astype(str).str.contains('💡 총계').any():
                                return ['background-color: #e0e2e6; font-weight: bold;'] * len(s)
                            return [''] * len(s)

                        fmt_dict = {prev_col_name: "{:,.0f}", curr_col_name: "{:,.0f}", "증감": "{:,.0f}", "증감률(%)": "{:,.1f}"}
                        st.data_editor(
                            center_style(df_show.style.format(fmt_dict).apply(highlight_map_total, axis=1)),
                            column_config={"선택": st.column_config.CheckboxColumn("선택", default=False)},
                            disabled=[c for c in df_show.columns if c != "선택"],
                            width="stretch", hide_index=True, key=editor_key
                        )
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
        def render_full_usage_report(usage_name, section_num, key_sfx_inner, db_key):
            st.markdown(
                f"""<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 10px;">
                <h4 style="margin: 0;">📈 {section_num}. 용도별 판매량 분석 : {usage_name}</h4></div>""",
                unsafe_allow_html=True
            )

            p_curr_act_all = pd.Series(dtype=float)
            p_prev_act_all = pd.Series(dtype=float)
            df_u_csv = pd.DataFrame()

            if not df_long_rpt.empty:
                df_u = df_long_rpt[df_long_rpt["그룹"] == usage_name]
                p_curr_act_all = df_u[(df_u["연"] == curr_year) & (df_u["계획/실적"] == "실적")].groupby("월")["값"].sum()
                p_prev_act_all = df_u[(df_u["연"] == prev_year) & (df_u["계획/실적"] == "실적")].groupby("월")["값"].sum()
            elif not df_csv_tab.empty and val_col in df_csv_tab.columns:
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

            if comp_mode == "YoY":
                if agg_mode == "누적 실적 (1월~당월)":
                    sum_act  = p_curr_act_all[p_curr_act_all.index <= curr_month].sum()
                    sum_prev = p_prev_act_all[p_prev_act_all.index <= prev_month].sum()
                    top_title = f"**■ 누적 실적 비교 ({curr_month}월 누적)**"
                else:
                    sum_act  = p_curr_act_all.get(curr_month, 0)
                    sum_prev = p_prev_act_all.get(prev_month, 0)
                    top_title = f"**■ 당월 실적 비교 ({curr_month}월 당월)**"
                prev_name   = f"{prev_year}년"
                curr_name   = f"{curr_year}년"
                diff_label  = "전년대비"
                vals_prev   = [p_prev_act_all.get(m, 0) for m in range(1, curr_month + 1)]
                prev_legend = f"{prev_year}년 실적"
            else:
                if agg_mode == "누적 실적 (1월~당월)":
                    sum_act  = p_curr_act_all[p_curr_act_all.index <= curr_month].sum()
                    sum_prev = p_prev_act_all[p_prev_act_all.index <= prev_month].sum()
                    top_title = f"**■ 누적 실적 비교 ({prev_month}월 누적 vs {curr_month}월 누적)**"
                else:
                    sum_act  = p_curr_act_all.get(curr_month, 0)
                    sum_prev = p_prev_act_all.get(prev_month, 0)
                    top_title = f"**■ 전월 실적 비교 ({prev_month}월 vs {curr_month}월)**"
                prev_name   = f"전월({prev_month}월)"
                curr_name   = f"당월({curr_month}월)"
                diff_label  = "전월대비"
                vals_prev = []
                for m in range(1, curr_month + 1):
                    if m > 1:
                        vals_prev.append(p_curr_act_all.get(m-1, 0))
                    else:
                        if not df_long_rpt.empty:
                            p_ly = df_long_rpt[
                                (df_long_rpt["그룹"] == usage_name) &
                                (df_long_rpt["연"] == curr_year-1) &
                                (df_long_rpt["계획/실적"] == "실적")
                            ].groupby("월")["값"].sum()
                            vals_prev.append(p_ly.get(12, 0))
                        elif not df_u_csv.empty:
                            p_ly = df_u_csv[df_u_csv["연_csv"] == curr_year-1].groupby("월_csv")[val_col].sum()
                            vals_prev.append(p_ly.get(12, 0))
                        else:
                            vals_prev.append(0)
                prev_legend = "전월 실적"

            diff_prev   = sum_act - sum_prev
            rate_prev   = (sum_act / sum_prev * 100) if sum_prev > 0 else 0
            sign_prev   = "+" if diff_prev > 0 else ""
            months_list = list(range(1, curr_month + 1))
            curr_legend = f"{curr_year}년 실적" if comp_mode == "YoY" else "당월 실적"
            desc_status = "감소" if diff_prev < 0 else "증가"

            st.markdown(
                f"""<div style="background-color: #f8f9fa; border-left: 5px solid #1e3a8a; padding: 15px;
                    margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                    <div style="font-size: 15px; color: #1e3a8a; font-weight: 700; line-height: 1.6;">
                        💡 [요약] 당해 실적: {sum_act:,.0f} {unit_str}<br>
                        {diff_label} <span style="color: {'#d32f2f' if diff_prev < 0 else '#1f77b4'};">
                        {abs(diff_prev):,.0f} {unit_str} {desc_status} ({sign_prev}{rate_prev:.1f}%)</span>
                    </div></div>""",
                unsafe_allow_html=True
            )

            vals_act    = [p_curr_act_all.get(m, 0) for m in months_list]
            overall_max = max(max([sum_prev, sum_act], default=0),
                              max(vals_act, default=0), max(vals_prev, default=0))
            yaxis_range = [0, overall_max * 1.1 if overall_max > 0 else 100]

            col_c, col_m = st.columns([1, 2.5])
            with col_c:
                st.markdown(top_title + f" <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                fig_c = go.Figure()
                fig_c.update_layout(margin=dict(t=30, b=20, l=40, r=10), height=420, showlegend=False)
                fig_c.update_yaxes(range=yaxis_range)
                fig_c.add_trace(go.Bar(
                    x=[f"{prev_name}<br>실적", f"{curr_name}<br>실적"], y=[sum_prev, sum_act],
                    marker_color=[COLOR_PREV, COLOR_ACT],
                    text=[f"{sum_prev:,.0f}", f"{sum_act:,.0f}"], textposition='auto', textfont=dict(size=14)
                ))
                st.plotly_chart(fig_c, width="stretch")

            with col_m:
                st.markdown(f"**■ 월별 실적 추이 (YoY)** <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                fig_m = go.Figure()
                fig_m.update_layout(
                    barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1),
                    margin=dict(t=30, b=20, l=40, r=10), height=420,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                fig_m.update_yaxes(range=yaxis_range)
                fig_m.add_trace(go.Bar(x=months_list, y=vals_prev, name=prev_legend, marker_color=COLOR_PREV,
                                       text=[f"{v:,.0f}" if v>0 else "" for v in vals_prev],
                                       textposition='auto', textfont=dict(size=11)))
                fig_m.add_trace(go.Bar(x=months_list, y=vals_act, name=curr_legend, marker_color=COLOR_ACT,
                                       text=[f"{v:,.0f}" if v>0 else "" for v in vals_act],
                                       textposition='auto', textfont=dict(size=11)))
                st.plotly_chart(fig_m, width="stretch")

            # 업종별 세부 분석
            if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                if "상품명" in df_csv_tab.columns:
                    csv_products2 = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                else:
                    csv_products2 = pd.Series([""] * len(df_csv_tab), index=df_csv_tab.index)

                if usage_name == "산업용":
                    df_sub_base = df_csv_tab[csv_products2 == "산업용"].copy()
                    grp_col = "업종"
                else:
                    valid_biz = ["냉난방용(업무)", "업무난방용", "주한미군"]
                    df_sub_base = df_csv_tab[csv_products2.isin(valid_biz)].copy()
                    if "업종분류" in df_sub_base.columns:
                        df_sub_base["업종"] = df_sub_base["업종분류"]
                    grp_col = "업종"

                mask_curr_sub = get_mask(df_sub_base, curr_year, curr_month, agg_mode)
                mask_prev_sub = get_mask(df_sub_base, prev_year, prev_month, agg_mode)

                if not df_sub_base.empty and grp_col in df_sub_base.columns:
                    curr_ind = df_sub_base[mask_curr_sub].groupby(grp_col, as_index=False)[val_col].sum().rename(columns={val_col: curr_name})
                    prev_ind = df_sub_base[mask_prev_sub].groupby(grp_col, as_index=False)[val_col].sum().rename(columns={val_col: prev_name})

                    ind_comp = pd.merge(prev_ind, curr_ind, on=grp_col, how="outer").fillna(0)
                    ind_comp = ind_comp.sort_values(curr_name, ascending=False).reset_index(drop=True)

                    if len(ind_comp) > 10:
                        top10 = ind_comp.iloc[:10].copy()
                        others = ind_comp.iloc[10:]
                        others_row = pd.DataFrame([{grp_col: "기타", prev_name: others[prev_name].sum(), curr_name: others[curr_name].sum()}])
                        ind_comp_plot = pd.concat([top10, others_row], ignore_index=True)
                    else:
                        ind_comp_plot = ind_comp.copy()

                    ind_comp_plot["증감절대값"] = abs(ind_comp_plot[curr_name] - ind_comp_plot[prev_name])
                    max_diff_idx = ind_comp_plot["증감절대값"].idxmax()
                    colors_act = [COLOR_ACT] * len(ind_comp_plot)
                    if pd.notna(max_diff_idx):
                        colors_act[int(max_diff_idx)] = "#d32f2f"

                    comp_title_suffix = "(당해연도 vs 전년도)" if comp_mode == "YoY" else "(당월 vs 전월)"
                    st.markdown(f"**■ 세부 업종별 판매량 비교 {comp_title_suffix}** <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)
                    fig_ind = go.Figure()
                    fig_ind.add_trace(go.Bar(x=ind_comp_plot[grp_col], y=ind_comp_plot[prev_name], name=prev_name,
                                            marker_color=COLOR_PREV,
                                            text=[f"{v:,.0f}" if v>0 else "" for v in ind_comp_plot[prev_name]],
                                            textposition='auto', textfont=dict(size=11)))
                    fig_ind.add_trace(go.Bar(x=ind_comp_plot[grp_col], y=ind_comp_plot[curr_name], name=curr_name,
                                            marker_color=colors_act,
                                            text=[f"{v:,.0f}" if v>0 else "" for v in ind_comp_plot[curr_name]],
                                            textposition='auto', textfont=dict(size=11)))
                    fig_ind.update_layout(barmode='group', xaxis_title="", yaxis_title=f"판매량({unit_str})",
                                         margin=dict(t=10, b=10, l=10, r=10), height=420,
                                         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                    st.plotly_chart(fig_ind, width="stretch")

                st.markdown("<hr style='border-top: 1px dashed #ccc; margin: 30px 0;'>", unsafe_allow_html=True)
                st.markdown(f"**🔍 {usage_name} 개별 고객 상세 차트** <span style='float:right; font-size:13px; font-weight:normal; color:gray;'>(단위: {unit_str})</span>", unsafe_allow_html=True)

                if not df_sub_base.empty and "고객명" in df_sub_base.columns:
                    c_curr_all = df_sub_base[mask_curr_sub].groupby(["고객명", grp_col], as_index=False)[val_col].sum().rename(columns={val_col: curr_name})
                    c_prev_all = df_sub_base[mask_prev_sub].groupby(["고객명", grp_col], as_index=False)[val_col].sum().rename(columns={val_col: prev_name})
                    grp_top = pd.merge(c_prev_all, c_curr_all, on=["고객명", grp_col], how="outer").fillna(0)
                    grp_top = grp_top.sort_values(curr_name, ascending=False).reset_index(drop=True)
                    grp_top = grp_top[(grp_top[curr_name] > 0) | (grp_top[prev_name] > 0)].reset_index(drop=True)

                    top_customers = [c for c in grp_top["고객명"] if "💡" not in str(c)]
                    sel_cust = st.selectbox(
                        f"상세 분석할 고객명을 선택하세요 ({usage_name})", ["선택 안함"] + top_customers,
                        key=f"sel_cust_{usage_name}_{key_sfx_inner}"
                    )

                    if sel_cust != "선택 안함":
                        # ✅ 개별 고객 조회도 반드시 df_csv_tab(단위변환 완료본)에서
                        c_data = df_csv_tab[df_csv_tab["고객명"] == sel_cust].copy()
                        mask_c_curr = get_mask(c_data, curr_year, curr_month, agg_mode)
                        mask_c_prev = get_mask(c_data, prev_year, prev_month, agg_mode)

                        sum_cur_c  = c_data[mask_c_curr][val_col].sum()
                        sum_prev_c = c_data[mask_c_prev][val_col].sum()

                        chart_title = (f"'{sel_cust}' 누적 사용량 ({curr_month}월 누적)"
                                       if agg_mode == "누적 실적 (1월~당월)"
                                       else f"'{sel_cust}' 당월 사용량 ({curr_month}월 당월)")
                        diff_val  = sum_cur_c - sum_prev_c
                        rate_val  = (sum_cur_c / sum_prev_c * 100) if sum_prev_c > 0 else 0
                        sign_str  = "+" if diff_val > 0 else ""
                        yoy_text  = f"{diff_label} 증감: {sign_str}{diff_val:,.0f} ({rate_val:.1f}%)"

                        cc1, cc2 = st.columns([1, 2])
                        with cc1:
                            fig_cust_cum = go.Figure()
                            fig_cust_cum.update_layout(title=chart_title, margin=dict(t=50, b=20, l=40, r=10), height=350)
                            fig_cust_cum.add_trace(go.Bar(
                                x=[prev_name, curr_name], y=[sum_prev_c, sum_cur_c],
                                marker_color=[COLOR_PREV, COLOR_ACT],
                                text=[f"{sum_prev_c:,.0f}", f"{sum_cur_c:,.0f}"], textposition='auto',
                                hovertemplate="%{x}: %{y:,.0f}<extra></extra>"
                            ))
                            fig_cust_cum.add_annotation(
                                x=0.5, y=1.05, xref="paper", yref="paper",
                                text=f"<b>{yoy_text}</b>", showarrow=False,
                                font=dict(size=13, color="#d32f2f" if diff_val < 0 else "#1f77b4"),
                                bgcolor="#f8f9fa", bordercolor="#d0d7e5", borderwidth=1, borderpad=4
                            )
                            st.plotly_chart(fig_cust_cum, width="stretch")

                        with cc2:
                            fig_cust_mon = go.Figure()
                            fig_cust_mon.update_layout(
                                title=f"'{sel_cust}' 월별 사용량 추이", barmode='group',
                                xaxis=dict(tickmode='linear', tick0=1, dtick=1),
                                margin=dict(t=50, b=20, l=40, r=10), height=350,
                                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                            )
                            cur_vals_c = [c_data[(c_data['연_csv']==curr_year) & (c_data['월_csv']==m)][val_col].sum() for m in months_list]
                            if comp_mode == "YoY":
                                prev_vals_c = [c_data[(c_data['연_csv']==prev_year) & (c_data['월_csv']==m)][val_col].sum() for m in months_list]
                            else:
                                prev_vals_c = []
                                for m in months_list:
                                    if m > 1:
                                        prev_vals_c.append(c_data[(c_data['연_csv']==curr_year) & (c_data['월_csv']==m-1)][val_col].sum())
                                    else:
                                        prev_vals_c.append(c_data[(c_data['연_csv']==curr_year-1) & (c_data['월_csv']==12)][val_col].sum())

                            fig_cust_mon.add_trace(go.Bar(
                                x=months_list, y=prev_vals_c, name=prev_legend, marker_color=COLOR_PREV,
                                text=[f"{v:,.0f}" if v>0 else "" for v in prev_vals_c],
                                textposition='auto', textfont=dict(size=11),
                                hovertemplate="%{x}월: %{y:,.0f}<extra></extra>"
                            ))
                            fig_cust_mon.add_trace(go.Bar(
                                x=months_list, y=cur_vals_c, name=curr_legend, marker_color=COLOR_ACT,
                                text=[f"{v:,.0f}" if v>0 else "" for v in cur_vals_c],
                                textposition='auto', textfont=dict(size=11),
                                hovertemplate="%{x}월: %{y:,.0f}<extra></extra>"
                            ))
                            st.plotly_chart(fig_cust_mon, width="stretch")

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

        st.html("""
            <button onclick="window.parent.print()" style="padding: 12px 20px; font-size: 16px;
                border-radius: 8px; background-color: #1e3a8a; color: white; border: none; cursor: pointer;
                width: 100%; font-weight: bold; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin: 2px;">
                🖨️ 현재 화면 전체를 PDF로 다운로드 (인쇄)
            </button>
        """)
