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
REPO_NAME = "Han11112222/quarterly-sales-report"
GITHUB_RAW_BASE = f"https://raw.githubusercontent.com/{REPO_NAME}/main"
COMMENT_DB_FILE = "report_comments_db.json"

COLOR_ACT  = "rgba(0, 150, 255, 1)"
COLOR_PREV = "rgba(190, 190, 190, 1)"

USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용":"가정용","개별난방용":"가정용","중앙난방용":"가정용","자가열전용":"가정용",
    "일반용":"영업용",
    "업무난방용":"업무용","냉방용":"업무용","주한미군":"업무용",
    "산업용":"산업용",
    "수송용(CNG)":"수송용","수송용(BIO)":"수송용",
    "열병합용":"열병합","열병합용1":"열병합","열병합용2":"열병합",
    "연료전지용":"연료전지","열전용설비용":"열전용설비용",
}

# ─────────────────────────────────────────────────────────
# 유틸 함수들
# ─────────────────────────────────────────────────────────
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
                repo.update_file(contents.path, "Update comments", content_string, contents.sha)
            except:
                repo.create_file(COMMENT_DB_FILE, "Create comments db", content_string)
    except Exception:
        pass

def clean_korean_finance_number(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(",", "").strip()
    if not s: return 0.0
    if s.endswith("-"): s = "-" + s[:-1]
    elif s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    s = re.sub(r"[^\d\.-]", "", s)
    try: return float(s)
    except: return 0.0

def center_style(styler):
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([
        dict(selector="th", props=[("text-align","center"),("vertical-align","middle"),
                                   ("background-color","#1e3a8a"),("color","#ffffff"),("font-weight","bold")]),
        dict(selector="thead th", props=[("background-color","#1e3a8a"),("color","#ffffff"),("font-weight","bold")]),
        dict(selector="tbody tr th", props=[("background-color","#1e3a8a"),("color","#ffffff"),("font-weight","bold")])
    ])
    return styler

def keyword_group(col: str) -> Optional[str]:
    c = str(col)
    if "열병합" in c: return "열병합"
    if "연료전지" in c: return "연료전지"
    if "수송용" in c: return "수송용"
    if "열전용" in c: return "열전용설비용"
    if c == "산업용": return "산업용"
    if c == "일반용": return "영업용"
    if any(k in c for k in ["취사용","난방용","자가열"]): return "가정용"
    if any(k in c for k in ["업무","냉방","주한미군"]): return "업무용"
    return None

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns: out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def make_long(plan_df, actual_df):
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)
    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연","월"]: continue
            group = USE_COL_TO_GROUP.get(col) or keyword_group(col)
            if group is None: continue
            base = df[["연","월"]].copy()
            base["그룹"] = group; base["용도"] = col
            base["계획/실적"] = label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)
    if not records: return pd.DataFrame(columns=["연","월","그룹","용도","계획/실적","값"])
    out = pd.concat(records, ignore_index=True).dropna(subset=["연","월"])
    out["연"] = out["연"].astype(int); out["월"] = out["월"].astype(int)
    return out

def load_all_sheets(excel_bytes):
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    needed = ["계획_부피","실적_부피","계획_열량","실적_열량"]
    return {n: xls.parse(n) for n in needed if n in xls.sheet_names}

def build_long_dict(sheets):
    out = {}
    if "계획_부피" in sheets and "실적_부피" in sheets:
        out["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if "계획_열량" in sheets and "실적_열량" in sheets:
        out["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return out

def load_safe_csv(file_bytes):
    for enc in ["utf-8-sig","cp949","utf-8","euc-kr"]:
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, thousands=',')
            df.columns = df.columns.str.strip()
            return df
        except Exception:
            pass
    return pd.DataFrame()

def get_coord_from_df(address, coord_df):
    if pd.isna(address) or not str(address).strip(): return None, None
    if not coord_df.empty and len(coord_df.columns) >= 3:
        clean_addr = re.sub(r'\(.*?\)', '', str(address)).split(',')[0].strip().replace(" ","")
        if clean_addr:
            addr_col, lat_col, lon_col = coord_df.columns[0], coord_df.columns[1], coord_df.columns[2]
            mask = coord_df[addr_col].astype(str).str.replace(" ","",regex=False).str.contains(re.escape(clean_addr), na=False)
            match = coord_df[mask]
            if not match.empty:
                try: return float(match.iloc[0][lat_col]), float(match.iloc[0][lon_col])
                except: pass
    return 35.8714 + random.uniform(-0.06,0.06), 128.6014 + random.uniform(-0.06,0.06)

def get_kakao_route(start_lon, start_lat, end_lon, end_lat):
    try:
        resp = requests.get(
            "https://apis-navi.kakaomobility.com/v1/directions",
            headers={"Authorization": "KakaoAK d9532fed7e56e09fe392c3482b915a20"},
            params={"origin": f"{start_lon},{start_lat}", "destination": f"{end_lon},{end_lat}", "priority": "RECOMMEND"},
            timeout=5
        )
        if resp.status_code == 200:
            coords = []
            routes = resp.json().get("routes", [])
            if routes:
                for section in routes[0]["sections"]:
                    for road in section["roads"]:
                        v = road["vertexes"]
                        coords.extend([[v[i], v[i+1]] for i in range(0, len(v), 2)])
            return coords
    except Exception:
        pass
    return None

def preprocess_csv(df_raw: pd.DataFrame) -> pd.DataFrame:
    """날짜 파싱만 수행. 단위 변환 없음."""
    if df_raw.empty: return df_raw.copy()
    df = df_raw.copy()
    date_col = next((c for c in ["청구년월","매출년월","년월","기준년월"] if c in df.columns), None)
    parsed = pd.Series([pd.NaT]*len(df), index=df.index)
    if date_col:
        try:
            p1 = pd.to_datetime(df[date_col], format="%b-%y", errors="coerce")
            p2 = pd.to_datetime(df[date_col], format="%Y%m", errors="coerce")
            p3 = pd.to_datetime(df[date_col], errors="coerce")
            parsed = p1.where(p1.notna(), p2).where(p1.notna()|p2.notna(), p3)
        except Exception:
            pass
    df["날짜_파싱"] = parsed.fillna(pd.to_datetime("2026-03-01"))
    df["연_csv"] = df["날짜_파싱"].dt.year.astype(int)
    df["월_csv"] = df["날짜_파싱"].dt.month.astype(int)
    return df

def get_unit_series(df: pd.DataFrame, unit_str: str) -> pd.Series:
    """단위 변환된 Series 반환. df 원본 불변."""
    if unit_str == "GJ" and "사용량(mj)" in df.columns:
        return pd.to_numeric(df["사용량(mj)"], errors="coerce").fillna(0.0) / 1000.0
    if unit_str == "천m³" and "사용량(m3)" in df.columns:
        return pd.to_numeric(df["사용량(m3)"], errors="coerce").fillna(0.0) / 1000.0
    for c in ["사용량(mj)","사용량(m3)"]:
        if c in df.columns:
            return pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return pd.Series([0.0]*len(df), index=df.index)

# ─────────────────────────────────────────────────────────
# ✅ get_mask: 루프 밖 전역 함수로 정의 (중복 정의 방지)
# ─────────────────────────────────────────────────────────
def get_mask(df, y, m, agg):
    if agg == "누적 실적 (1월~당월)":
        return (df["연_csv"] == y) & (df["월_csv"] <= m)
    return (df["연_csv"] == y) & (df["월_csv"] == m)

# ─────────────────────────────────────────────────────────
# GitHub API로 CSV 다운로드
# ─────────────────────────────────────────────────────────
def load_csvs_via_github_api() -> pd.DataFrame:
    """
    GitHub Raw URL로 파일명을 직접 생성해서 다운로드.
    API 호출 없이 public repo에서 직접 다운로드하므로 토큰 불필요.
    """
    # 2025년 1~12월, 2026년 1~12월 파일명 생성
    filenames = []
    for y in [2025, 2026]:
        for m in range(1, 13):
            filenames.append(f"가정용외_{y}{m:02d}.csv")

    csv_list, loaded_names = [], []
    for fname in filenames:
        url = f"{GITHUB_RAW_BASE}/{fname}"
        try:
            r = requests.get(url, timeout=15)
            if r.status_code == 200:
                df = load_safe_csv(r.content)
                if not df.empty:
                    csv_list.append(df)
                    loaded_names.append(fname)
        except Exception:
            pass

    if csv_list:
        st.session_state["github_csv_loaded"] = loaded_names
        return pd.concat(csv_list, ignore_index=True)

    st.session_state["github_csv_error"] = "GitHub Raw URL로 CSV를 불러오지 못했습니다."
    return pd.DataFrame()

# ─────────────────────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────────────────────
st.title("📊 대용량 수요처 이상 감지 대시보드")

with st.sidebar:
    st.header("📂 데이터 및 설정")

    st.subheader("1. 판매량 데이터 (요약/엑셀)")
    src_sales = st.radio("판매량 데이터 소스", ["레포 파일 사용","엑셀 업로드(.xlsx)"], index=0, key="rpt_sales_src")
    excel_bytes = None
    if src_sales == "엑셀 업로드(.xlsx)":
        up_sales = st.file_uploader("판매량(계획_실적).xlsx 형식", type=["xlsx"], key="rpt_sales_uploader")
        if up_sales: excel_bytes = up_sales.getvalue()
    else:
        p = Path(__file__).parent / DEFAULT_SALES_XLSX
        if p.exists(): excel_bytes = p.read_bytes()

    st.markdown("---")
    st.subheader("2. 업종별 데이터 (상세/CSV)")
    src_csv = st.radio("업종별 데이터 소스", ["레포 파일 사용","CSV 업로드(.csv)"], index=0, key="csv_src")
    if src_csv == "CSV 업로드(.csv)":
        up_csvs = st.file_uploader("가정용외_*.csv (다중 업로드)", type=["csv"], accept_multiple_files=True, key="csv_uploader")
        if up_csvs:
            dfs = [load_safe_csv(f.getvalue()) for f in up_csvs]
            dfs = [d for d in dfs if not d.empty]
            if dfs: st.session_state['merged_csv_df'] = pd.concat(dfs, ignore_index=True)
        else:
            st.session_state.pop('merged_csv_df', None)
    else:
        if "github_csv_loaded" in st.session_state:
            names = st.session_state["github_csv_loaded"]
            st.success(f"✅ CSV {len(names)}개 로드 완료")
            with st.expander("로드된 파일 목록"):
                for n in names: st.write(f"- {n}")
        if "github_csv_error" in st.session_state:
            st.error(f"로드 오류: {st.session_state['github_csv_error']}")

    st.markdown("---")
    st.subheader("🗺️ 3. 지도 위경도 데이터 (CSV)")
    src_coord = st.radio("위경도 데이터 소스", ["레포 파일(깃허브) 사용","CSV 업로드(.csv)"], index=0, key="coord_src")
    coord_df = pd.DataFrame()
    if src_coord == "CSV 업로드(.csv)":
        up_coord = st.file_uploader("address_with_latlon.csv", type=["csv"], key="coord_uploader")
        if up_coord: coord_df = load_safe_csv(up_coord.getvalue())
    else:
        cp = Path(__file__).parent / "address_with_latlon.csv"
        if cp.exists():
            coord_df = load_safe_csv(cp.read_bytes())
        else:
            try:
                r = requests.get(f"{GITHUB_RAW_BASE}/address_with_latlon.csv", timeout=5)
                if r.status_code == 200: coord_df = load_safe_csv(r.content)
            except: pass

# ─────────────────────────────────────────────────────────
# 본문: 데이터 로드
# ─────────────────────────────────────────────────────────
long_dict_rpt: Dict[str, pd.DataFrame] = {}
if excel_bytes:
    long_dict_rpt = build_long_dict(load_all_sheets(excel_bytes))

# CSV 원본 로드 (단위 변환 없이)
df_csv_raw = pd.DataFrame()
if src_csv == "레포 파일 사용":
    # ✅ 항상 GitHub API로 로드 (glob은 Streamlit Cloud에서 불안정)
    df_csv_raw = load_csvs_via_github_api()

if df_csv_raw.empty and 'merged_csv_df' in st.session_state:
    df_csv_raw = st.session_state['merged_csv_df'].copy()

# 숫자 정제
if not df_csv_raw.empty:
    for col in ["사용량(mj)","사용량(m3)"]:
        if col in df_csv_raw.columns:
            df_csv_raw[col] = df_csv_raw[col].apply(clean_korean_finance_number)

# 날짜 파싱 (한 번만)
df_csv_parsed = preprocess_csv(df_csv_raw) if not df_csv_raw.empty else pd.DataFrame()

comments_db = load_comments_db()
UNIT_VAL_COL = "사용량_단위변환"

# ─────────────────────────────────────────────────────────
# 탭
# ─────────────────────────────────────────────────────────
rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

for tab_idx, rpt_tab in enumerate(rpt_tabs):
    with rpt_tab:
        if tab_idx == 0:
            df_long_rpt = long_dict_rpt.get("열량", pd.DataFrame())
            unit_str = "GJ"; key_sfx = "_gj"
        else:
            df_long_rpt = long_dict_rpt.get("부피", pd.DataFrame())
            unit_str = "천m³"; key_sfx = "_vol"

        # 탭별 단위변환 (원본 컬럼 불변, 새 컬럼 추가)
        if not df_csv_parsed.empty:
            df_csv_tab = df_csv_parsed.copy()
            df_csv_tab[UNIT_VAL_COL] = get_unit_series(df_csv_tab, unit_str)
        else:
            df_csv_tab = pd.DataFrame()
        val_col = UNIT_VAL_COL

        # 기준 일자 설정
        st.markdown("#### 📅 기준 일자 설정")
        years_available = [2024, 2025, 2026]
        default_y_index = len(years_available) - 1
        default_m_index = 2

        if not df_long_rpt.empty:
            years_available = sorted(df_long_rpt["연"].unique().tolist())
            actual_data = df_long_rpt[(df_long_rpt["계획/실적"]=="실적") & (df_long_rpt["값"]>0)]
            if not actual_data.empty:
                my = int(actual_data["연"].max())
                mm = int(actual_data[actual_data["연"]==my]["월"].max())
                default_y_index = years_available.index(my) if my in years_available else len(years_available)-1
                default_m_index = mm - 1

        if not df_csv_tab.empty and "날짜_파싱" in df_csv_tab.columns:
            csv_max = df_csv_tab["날짜_파싱"].max()
            if pd.notna(csv_max):
                y = int(csv_max.year)
                if y not in years_available: years_available = sorted(set(years_available)|{y})
                default_y_index = years_available.index(y)
                default_m_index = int(csv_max.month) - 1

        c_y, c_m, c_agg, _ = st.columns([1,1,2,1])
        with c_y:
            sel_year_rpt = st.selectbox("기준 연도", years_available, index=default_y_index, key=f"rpt_yr{key_sfx}")
        with c_m:
            sel_month_str = st.selectbox("기준 월", [f"{m}월" for m in range(1,13)], index=default_m_index, key=f"rpt_mo{key_sfx}")
        with c_agg:
            agg_mode = st.radio("집계 기준", ["당월 실적","누적 실적 (1월~당월)"], index=0, horizontal=True, key=f"agg_mode_{key_sfx}")

        max_month  = int(sel_month_str.replace("월",""))
        curr_year  = sel_year_rpt
        curr_month = max_month

        st.markdown("<hr style='margin:10px 0 30px 0;'>", unsafe_allow_html=True)

        # 비교 기준 설정
        map_c1, map_c2, map_c3 = st.columns([1,1,1])
        with map_c1:
            map_usage = st.radio("📍 지도에 표시할 용도 선택", ["산업용","업무용"], index=0, horizontal=True, key=f"map_radio_{key_sfx}")
        with map_c2:
            comp_mode = st.radio("📍 비교 기준", ["YoY","전월대비"], index=0, horizontal=True, key=f"comp_mode_{key_sfx}")
        with map_c3:
            map_style_ui = st.radio("📍 지도 배경 테마", ["다크 모드 (기본)","일반 도로 지도"], index=0, horizontal=True, key=f"map_style_{key_sfx}")

        deck_map_style = "dark" if map_style_ui == "다크 모드 (기본)" else "road"

        if comp_mode == "YoY":
            prev_year = curr_year - 1; prev_month = curr_month
        else:
            prev_year = curr_year if curr_month > 1 else curr_year - 1
            prev_month = curr_month - 1 if curr_month > 1 else 12

        # ── 1. 이상 감지 지도 ──
        st.markdown(
            f"### 🗺️ 1. 대용량 수요처 이상 감지 모니터링 지도 "
            f"<span style='float:right;font-size:13px;font-weight:normal;color:gray;'>(단위:{unit_str})</span>",
            unsafe_allow_html=True)
        st.caption("※ YoY 기준 5% 이상 사용량이 하락한 업체를 지도에 마커로 표시하여 현장 방문을 유도합니다.")
        st.markdown("""
        <div style='background-color:#f1f3f5;padding:12px;border-radius:6px;margin-bottom:15px;font-size:14px;'>
            <b>💡 지도 마커(알람) 3단계 구분 안내</b><br>
            • <b>심각 (20% 이상 하락)</b> : 가장 크고 진한 색상의 마커<br>
            • <b>경계 (10% 이상 하락)</b> : 중간 크기와 중간 농도의 마커<br>
            • <b>주의 (5% 이상 하락)</b> : 작고 연한 색상의 마커<br>
            <span style='font-size:12px;color:#555;'>※ 산업용은 붉은색(🔴), 업무용은 푸른색(🔵) 계열로 표시됩니다.</span>
        </div>""", unsafe_allow_html=True)

        if (not df_csv_tab.empty and "도로명주소" in df_csv_tab.columns
                and "고객명" in df_csv_tab.columns and val_col in df_csv_tab.columns
                and "용도" in df_csv_tab.columns):

            if map_usage == "산업용":
                df_map = df_csv_tab[df_csv_tab["용도"]=="산업용"].copy()
            else:
                if "상품명" in df_csv_tab.columns:
                    ps = df_csv_tab["상품명"].astype(str).str.replace(r"\s+","",regex=True)
                    df_map = df_csv_tab[(df_csv_tab["용도"]=="업무용")|ps.isin(["냉난방용(업무)","업무난방용","주한미군"])].copy()
                else:
                    df_map = df_csv_tab[df_csv_tab["용도"]=="업무용"].copy()
            df_map["용도_태그"] = f"[{map_usage}]"

            m_curr = get_mask(df_map, curr_year, curr_month, agg_mode)
            m_prev = get_mask(df_map, prev_year, prev_month, agg_mode)
            map_curr = df_map[m_curr].groupby(["고객명","도로명주소","용도_태그"],as_index=False)[val_col].sum().rename(columns={val_col:"당해년도"})
            map_prev = df_map[m_prev].groupby(["고객명","도로명주소","용도_태그"],as_index=False)[val_col].sum().rename(columns={val_col:"전년도"})

            if not map_curr.empty and not map_prev.empty:
                merged = pd.merge(map_curr, map_prev, on=["고객명","도로명주소","용도_태그"], how="inner").fillna(0)
                merged["증감률(%)"] = np.where(merged["전년도"]>0, (merged["당해년도"]-merged["전년도"])/merged["전년도"]*100, 0)
                alarm_df = merged[merged["증감률(%)"]<=-5].copy()

                if alarm_df.empty:
                    st.success(f"✅ 선택한 기간 내 YoY 5% 이상 하락한 {map_usage} 리스크 업체가 없습니다.")
                else:
                    st.warning(f"🚨 총 **{len(alarm_df)}**개의 {map_usage} 업체에서 5% 이상 하락 신호가 감지되었습니다.")
                    alarm_df["감소량"] = alarm_df["전년도"] - alarm_df["당해년도"]
                    alarm_df = alarm_df.sort_values("감소량", ascending=False).head(100).reset_index(drop=True)
                    alarm_df["증감"] = alarm_df["당해년도"] - alarm_df["전년도"]

                    lats,lons,tooltips,colors,radiuses = [],[],[],[],[]
                    for _,row in alarm_df.iterrows():
                        lat,lon = get_coord_from_df(row['도로명주소'], coord_df)
                        lats.append(lat); lons.append(lon)
                        rate = row['증감률(%)']
                        if map_usage == "산업용":
                            if rate<=-20:   lv,c,r = "심각",[180,0,0,255],150
                            elif rate<=-10: lv,c,r = "경계",[255,80,80,200],100
                            else:           lv,c,r = "주의",[255,150,150,200],80
                        else:
                            if rate<=-20:   lv,c,r = "심각",[0,0,180,255],150
                            elif rate<=-10: lv,c,r = "경계",[80,150,255,200],100
                            else:           lv,c,r = "주의",[120,180,255,200],80
                        colors.append(c); radiuses.append(r)
                        tooltips.append(
                            f"<b>{row['용도_태그']} {row['고객명']} <span style='color:red;'>[{lv}]</span></b><br/>"
                            f"전년/전월:{row['전년도']:,.0f} / 당해:{row['당해년도']:,.0f}<br/>"
                            f"증감률:<span style='color:red;font-weight:bold;'>{row['증감률(%)']:.1f}%</span><br/>"
                            f"<span style='font-size:0.8em;color:gray;'>{row['도로명주소']}</span>")

                    alarm_df['lat']=lats; alarm_df['lon']=lons
                    alarm_df['tooltip']=tooltips; alarm_df['color']=colors; alarm_df['radius']=radiuses
                    alarm_df = alarm_df.dropna(subset=['lat','lon'])

                    if not alarm_df.empty:
                        editor_key = f"editor_{map_usage}_{key_sfx}"
                        selected_indices = []
                        if editor_key in st.session_state:
                            for ks,rv in st.session_state[editor_key].get("edited_rows",{}).items():
                                if rv.get("선택",False):
                                    ki = int(ks)
                                    if ki < len(alarm_df): selected_indices.append(ki)

                        map_df = alarm_df.iloc[selected_indices].copy() if selected_indices else alarm_df.copy()
                        layers = [pdk.Layer("ScatterplotLayer", data=map_df,
                            get_position='[lon,lat]', get_color='color', get_radius='radius',
                            pickable=True, opacity=0.6, filled=True, stroked=True,
                            get_line_color=[255,255,255,200], line_width_min_pixels=1, radius_max_pixels=40)]

                        start_lat,start_lon = 35.8660194, 128.5332943
                        if selected_indices:
                            draw_route = st.button("🚗 선택 업체 최적 동선(실제 도로) 그리기", width="stretch", key=f"draw_route_btn_{key_sfx}")
                            layers.append(pdk.Layer("ScatterplotLayer",
                                data=pd.DataFrame([{"lon":start_lon,"lat":start_lat,
                                    "tooltip":"<b>🏢 대성에너지 서부지사 (출발지)</b>",
                                    "color":[255,193,7,255],"radius":150}]),
                                get_position='[lon,lat]', get_color='color', get_radius='radius',
                                pickable=True, opacity=1.0, filled=True, stroked=True,
                                get_line_color=[0,0,0,255], line_width_min_pixels=2, radius_max_pixels=15))
                            if draw_route:
                                with st.spinner("카카오 모빌리티 API를 통해 최적 도로 경로를 탐색 중입니다..."):
                                    unvisited = map_df[['lon','lat','고객명']].to_dict('records')
                                    cl,cln = start_lat, start_lon
                                    ordered = []
                                    while unvisited:
                                        near = min(unvisited, key=lambda p: (p['lat']-cl)**2+(p['lon']-cln)**2)
                                        ordered.append(near); cl,cln=near['lat'],near['lon']; unvisited.remove(near)
                                    full_coords=[]; cur=[start_lon,start_lat]
                                    for stop in ordered:
                                        tgt=[stop['lon'],stop['lat']]
                                        seg=get_kakao_route(cur[0],cur[1],tgt[0],tgt[1])
                                        full_coords.extend(seg if seg else [cur,tgt]); cur=tgt
                                    if full_coords:
                                        layers.append(pdk.Layer("PathLayer",
                                            data=pd.DataFrame([{"path":full_coords,"color":[46,204,113,255]}]),
                                            get_path="path", get_color="color", width_scale=20, width_min_pixels=3, get_width=5))
                                        st.success("✨ 최적 도로 경로가 지도에 표시되었습니다!")

                        st.pydeck_chart(pdk.Deck(
                            map_style=deck_map_style, layers=layers,
                            initial_view_state=pdk.ViewState(
                                latitude=start_lat if selected_indices else alarm_df['lat'].mean(),
                                longitude=start_lon if selected_indices else alarm_df['lon'].mean(),
                                zoom=11, pitch=40),
                            tooltip={"html":"{tooltip}","style":{"backgroundColor":"white","color":"black"}}))

                        prev_cn = "전년도" if comp_mode=="YoY" else "전월"
                        curr_cn = "당해년도" if comp_mode=="YoY" else "당월"
                        df_show = alarm_df[['용도_태그','고객명','도로명주소','전년도','당해년도','증감','증감률(%)']].copy()
                        df_show = df_show.rename(columns={"전년도":prev_cn,"당해년도":curr_cn})
                        df_show.insert(0,"No.",range(1,len(df_show)+1))
                        df_show.insert(0,"선택",False)
                        df_show["비고"] = np.where(df_show["증감률(%)"]<=-99.9,"폐업의심","")
                        sp=df_show[prev_cn].sum(); sc=df_show[curr_cn].sum()
                        sr=((sc-sp)/sp*100) if sp>0 else 0
                        df_show = pd.concat([df_show, pd.DataFrame([{
                            "선택":False,"No.":"","용도_태그":"💡 총계","고객명":"","도로명주소":"",
                            prev_cn:sp, curr_cn:sc, "증감":sc-sp, "증감률(%)":sr, "비고":""}])], ignore_index=True)

                        def _hl(s):
                            return (['background-color:#e0e2e6;font-weight:bold;']*len(s)
                                    if s.astype(str).str.contains('💡 총계').any() else ['']*len(s))

                        fmt = {prev_cn:"{:,.0f}",curr_cn:"{:,.0f}","증감":"{:,.0f}","증감률(%)":"{:,.1f}"}
                        st.markdown(
                            f"<br><b>📋 지도 표기 업체 요약표</b> "
                            f"<span style='font-size:13px;color:#d32f2f;margin-left:10px;'>✅ 표 좌측 [선택] 체크 시 상단 지도에 선택 업체와 출발지가 뜹니다.</span> "
                            f"<span style='float:right;font-size:13px;font-weight:normal;color:gray;'>(단위:{unit_str})</span>",
                            unsafe_allow_html=True)
                        st.data_editor(
                            center_style(df_show.style.format(fmt).apply(_hl,axis=1)),
                            column_config={"선택":st.column_config.CheckboxColumn("선택",default=False)},
                            disabled=[c for c in df_show.columns if c!="선택"],
                            width="stretch", hide_index=True, key=editor_key)
                    else:
                        st.error("매핑된 위경도 좌표가 없어 지도를 표시할 수 없습니다.")
            else:
                st.info("비교할 과거 또는 당해 연도 데이터가 없습니다.")
        else:
            st.info("데이터에 '도로명주소', '고객명', '용도' 컬럼이 없거나 데이터가 부족합니다.")

        st.markdown("<hr style='border-top:2px solid #1e3a8a;margin:50px 0 20px 0;'>", unsafe_allow_html=True)

        # ── 2,3. 용도별 분석 함수 ──
        def render_usage(usage_name, section_num, ks):
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:15px;margin-bottom:10px;'>"
                f"<h4 style='margin:0;'>📈 {section_num}. 용도별 판매량 분석 : {usage_name}</h4></div>",
                unsafe_allow_html=True)

            p_curr = pd.Series(dtype=float)
            p_prev = pd.Series(dtype=float)
            df_u   = pd.DataFrame()

            if not df_long_rpt.empty:
                grp = df_long_rpt[df_long_rpt["그룹"]==usage_name]
                p_curr = grp[(grp["연"]==curr_year)&(grp["계획/실적"]=="실적")].groupby("월")["값"].sum()
                p_prev = grp[(grp["연"]==prev_year)&(grp["계획/실적"]=="실적")].groupby("월")["값"].sum()
            elif not df_csv_tab.empty and val_col in df_csv_tab.columns:
                ps2 = df_csv_tab["상품명"].astype(str).str.replace(r"\s+","",regex=True) if "상품명" in df_csv_tab.columns else pd.Series([""]*len(df_csv_tab),index=df_csv_tab.index)
                if usage_name=="산업용":
                    df_u = df_csv_tab[(df_csv_tab["용도"]=="산업용") | (ps2=="산업용")].copy()
                else:
                    mask_biz = (df_csv_tab["용도"]=="업무용") | ps2.isin(["냉난방용(업무)","업무난방용","주한미군"])
                    df_u = df_csv_tab[mask_biz].copy()
                p_curr = df_u[df_u["연_csv"]==curr_year].groupby("월_csv")[val_col].sum()
                p_prev = df_u[df_u["연_csv"]==prev_year].groupby("월_csv")[val_col].sum()

            if comp_mode=="YoY":
                if agg_mode=="누적 실적 (1월~당월)":
                    s_act=p_curr[p_curr.index<=curr_month].sum(); s_prev=p_prev[p_prev.index<=prev_month].sum()
                    ttl=f"**■ 누적 실적 비교 ({curr_month}월 누적)**"
                else:
                    s_act=p_curr.get(curr_month,0); s_prev=p_prev.get(prev_month,0)
                    ttl=f"**■ 당월 실적 비교 ({curr_month}월 당월)**"
                pn=f"{prev_year}년"; cn=f"{curr_year}년"; dl="전년대비"
                vp=[p_prev.get(m,0) for m in range(1,curr_month+1)]; pl=f"{prev_year}년 실적"
            else:
                if agg_mode=="누적 실적 (1월~당월)":
                    s_act=p_curr[p_curr.index<=curr_month].sum(); s_prev=p_prev[p_prev.index<=prev_month].sum()
                    ttl=f"**■ 누적 실적 비교 ({prev_month}월 누적 vs {curr_month}월 누적)**"
                else:
                    s_act=p_curr.get(curr_month,0); s_prev=p_prev.get(prev_month,0)
                    ttl=f"**■ 전월 실적 비교 ({prev_month}월 vs {curr_month}월)**"
                pn=f"전월({prev_month}월)"; cn=f"당월({curr_month}월)"; dl="전월대비"
                vp=[]
                for m in range(1,curr_month+1):
                    if m>1: vp.append(p_curr.get(m-1,0))
                    else:
                        if not df_long_rpt.empty:
                            ly=df_long_rpt[(df_long_rpt["그룹"]==usage_name)&(df_long_rpt["연"]==curr_year-1)&(df_long_rpt["계획/실적"]=="실적")].groupby("월")["값"].sum()
                            vp.append(ly.get(12,0))
                        elif not df_u.empty:
                            ly2=df_u[df_u["연_csv"]==curr_year-1].groupby("월_csv")[val_col].sum()
                            vp.append(ly2.get(12,0))
                        else: vp.append(0)
                pl="전월 실적"

            diff=s_act-s_prev; rate=(s_act/s_prev*100) if s_prev>0 else 0
            sign="+" if diff>0 else ""; ml=list(range(1,curr_month+1))
            cl2=f"{curr_year}년 실적" if comp_mode=="YoY" else "당월 실적"
            ds="감소" if diff<0 else "증가"

            st.markdown(
                f"<div style='background-color:#f8f9fa;border-left:5px solid #1e3a8a;padding:15px;"
                f"margin-bottom:20px;border-radius:4px;'>"
                f"<div style='font-size:15px;color:#1e3a8a;font-weight:700;line-height:1.6;'>"
                f"💡 [요약] 당해 실적: {s_act:,.0f} {unit_str}<br>"
                f"{dl} <span style='color:{'#d32f2f' if diff<0 else '#1f77b4'};'>"
                f"{abs(diff):,.0f} {unit_str} {ds} ({sign}{rate:.1f}%)</span></div></div>",
                unsafe_allow_html=True)

            va=[p_curr.get(m,0) for m in ml]
            ym=max(max([s_prev,s_act],default=0),max(va,default=0),max(vp,default=0))
            yr=[0,ym*1.1 if ym>0 else 100]

            c1,c2=st.columns([1,2.5])
            with c1:
                st.markdown(ttl+f" <span style='float:right;font-size:13px;font-weight:normal;color:gray;'>(단위:{unit_str})</span>",unsafe_allow_html=True)
                fc=go.Figure()
                fc.update_layout(margin=dict(t=30,b=20,l=40,r=10),height=420,showlegend=False)
                fc.update_yaxes(range=yr)
                fc.add_trace(go.Bar(x=[f"{pn}<br>실적",f"{cn}<br>실적"],y=[s_prev,s_act],
                    marker_color=[COLOR_PREV,COLOR_ACT],text=[f"{s_prev:,.0f}",f"{s_act:,.0f}"],
                    textposition='auto',textfont=dict(size=14)))
                st.plotly_chart(fc, key=f"fc_{usage_name}_{ks}", width="stretch")
            with c2:
                st.markdown(f"**■ 월별 실적 추이** <span style='float:right;font-size:13px;font-weight:normal;color:gray;'>(단위:{unit_str})</span>",unsafe_allow_html=True)
                fm=go.Figure()
                fm.update_layout(barmode='group',xaxis=dict(tickmode='linear',tick0=1,dtick=1),
                    margin=dict(t=30,b=20,l=40,r=10),height=420,
                    legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
                fm.update_yaxes(range=yr)
                fm.add_trace(go.Bar(x=ml,y=vp,name=pl,marker_color=COLOR_PREV,
                    text=[f"{v:,.0f}" if v>0 else "" for v in vp],textposition='auto',textfont=dict(size=11)))
                fm.add_trace(go.Bar(x=ml,y=va,name=cl2,marker_color=COLOR_ACT,
                    text=[f"{v:,.0f}" if v>0 else "" for v in va],textposition='auto',textfont=dict(size=11)))
                st.plotly_chart(fm, key=f"fm_{usage_name}_{ks}", width="stretch")

            if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                ps3 = df_csv_tab["상품명"].astype(str).str.replace(r"\s+","",regex=True) if "상품명" in df_csv_tab.columns else pd.Series([""]*len(df_csv_tab),index=df_csv_tab.index)
                if usage_name=="산업용":
                    dsb=df_csv_tab[(df_csv_tab["용도"]=="산업용") | (ps3=="산업용")].copy(); gc="업종"
                else:
                    mask_biz2 = (df_csv_tab["용도"]=="업무용") | ps3.isin(["냉난방용(업무)","업무난방용","주한미군"])
                    dsb=df_csv_tab[mask_biz2].copy()
                    if "업종분류" in dsb.columns: dsb["업종"]=dsb["업종분류"]
                    gc="업종"

                mcs=get_mask(dsb,curr_year,curr_month,agg_mode)
                mps=get_mask(dsb,prev_year,prev_month,agg_mode)

                if not dsb.empty and gc in dsb.columns:
                    ci=dsb[mcs].groupby(gc,as_index=False)[val_col].sum().rename(columns={val_col:cn})
                    pi=dsb[mps].groupby(gc,as_index=False)[val_col].sum().rename(columns={val_col:pn})
                    ic=pd.merge(pi,ci,on=gc,how="outer").fillna(0).sort_values(cn,ascending=False).reset_index(drop=True)
                    if len(ic)>10:
                        t10=ic.iloc[:10].copy(); oth=ic.iloc[10:]
                        ic=pd.concat([t10,pd.DataFrame([{gc:"기타",pn:oth[pn].sum(),cn:oth[cn].sum()}])],ignore_index=True)
                    ic["d"]=abs(ic[cn]-ic[pn]); mx=ic["d"].idxmax()
                    ca=[COLOR_ACT]*len(ic)
                    if pd.notna(mx): ca[int(mx)]="#d32f2f"
                    st.markdown(f"**■ 세부 업종별 판매량 비교** <span style='float:right;font-size:13px;font-weight:normal;color:gray;'>(단위:{unit_str})</span>",unsafe_allow_html=True)
                    fi=go.Figure()
                    fi.add_trace(go.Bar(x=ic[gc],y=ic[pn],name=pn,marker_color=COLOR_PREV,
                        text=[f"{v:,.0f}" if v>0 else "" for v in ic[pn]],textposition='auto',textfont=dict(size=11)))
                    fi.add_trace(go.Bar(x=ic[gc],y=ic[cn],name=cn,marker_color=ca,
                        text=[f"{v:,.0f}" if v>0 else "" for v in ic[cn]],textposition='auto',textfont=dict(size=11)))
                    fi.update_layout(barmode='group',xaxis_title="",yaxis_title=f"판매량({unit_str})",
                        margin=dict(t=10,b=10,l=10,r=10),height=420,
                        legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
                    st.plotly_chart(fi, key=f"fi_{usage_name}_{ks}", width="stretch")

                st.markdown("<hr style='border-top:1px dashed #ccc;margin:30px 0;'>",unsafe_allow_html=True)
                st.markdown(f"**🔍 {usage_name} 개별 고객 상세 차트** <span style='float:right;font-size:13px;font-weight:normal;color:gray;'>(단위:{unit_str})</span>",unsafe_allow_html=True)

                if not dsb.empty and "고객명" in dsb.columns:
                    cc=dsb[mcs].groupby(["고객명",gc],as_index=False)[val_col].sum().rename(columns={val_col:cn})
                    pc2=dsb[mps].groupby(["고객명",gc],as_index=False)[val_col].sum().rename(columns={val_col:pn})
                    gt=pd.merge(pc2,cc,on=["고객명",gc],how="outer").fillna(0)
                    gt=gt.sort_values(cn,ascending=False).reset_index(drop=True)
                    gt=gt[(gt[cn]>0)|(gt[pn]>0)].reset_index(drop=True)
                    tops=[c for c in gt["고객명"] if "💡" not in str(c)]
                    sel=st.selectbox(f"상세 분석할 고객명 ({usage_name})",["선택 안함"]+tops,key=f"sel_{usage_name}_{ks}")

                    if sel!="선택 안함":
                        cd=df_csv_tab[df_csv_tab["고객명"]==sel].copy()
                        sc2=cd[get_mask(cd,curr_year,curr_month,agg_mode)][val_col].sum()
                        sp2=cd[get_mask(cd,prev_year,prev_month,agg_mode)][val_col].sum()
                        ct=(f"'{sel}' 누적 사용량 ({curr_month}월 누적)" if agg_mode=="누적 실적 (1월~당월)"
                            else f"'{sel}' 당월 사용량 ({curr_month}월 당월)")
                        dv=sc2-sp2; rv=(sc2/sp2*100) if sp2>0 else 0
                        yt=f"{dl} 증감: {'+' if dv>0 else ''}{dv:,.0f} ({rv:.1f}%)"

                        x1,x2=st.columns([1,2])
                        with x1:
                            fg=go.Figure()
                            fg.update_layout(title=ct,margin=dict(t=50,b=20,l=40,r=10),height=350)
                            fg.add_trace(go.Bar(x=[pn,cn],y=[sp2,sc2],marker_color=[COLOR_PREV,COLOR_ACT],
                                text=[f"{sp2:,.0f}",f"{sc2:,.0f}"],textposition='auto',
                                hovertemplate="%{x}: %{y:,.0f}<extra></extra>"))
                            fg.add_annotation(x=0.5,y=1.05,xref="paper",yref="paper",text=f"<b>{yt}</b>",
                                showarrow=False,font=dict(size=13,color="#d32f2f" if dv<0 else "#1f77b4"),
                                bgcolor="#f8f9fa",bordercolor="#d0d7e5",borderwidth=1,borderpad=4)
                            st.plotly_chart(fg, key=f"fg_{usage_name}_{ks}_{sel}", width="stretch")
                        with x2:
                            fm2=go.Figure()
                            fm2.update_layout(title=f"'{sel}' 월별 사용량 추이",barmode='group',
                                xaxis=dict(tickmode='linear',tick0=1,dtick=1),
                                margin=dict(t=50,b=20,l=40,r=10),height=350,
                                legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
                            cvc=[cd[(cd['연_csv']==curr_year)&(cd['월_csv']==m)][val_col].sum() for m in ml]
                            if comp_mode=="YoY":
                                pvc=[cd[(cd['연_csv']==prev_year)&(cd['월_csv']==m)][val_col].sum() for m in ml]
                            else:
                                pvc=[]
                                for m in ml:
                                    if m>1: pvc.append(cd[(cd['연_csv']==curr_year)&(cd['월_csv']==m-1)][val_col].sum())
                                    else:   pvc.append(cd[(cd['연_csv']==curr_year-1)&(cd['월_csv']==12)][val_col].sum())
                            fm2.add_trace(go.Bar(x=ml,y=pvc,name=pl,marker_color=COLOR_PREV,
                                text=[f"{v:,.0f}" if v>0 else "" for v in pvc],textposition='auto',textfont=dict(size=11),
                                hovertemplate="%{x}월: %{y:,.0f}<extra></extra>"))
                            fm2.add_trace(go.Bar(x=ml,y=cvc,name=cl2,marker_color=COLOR_ACT,
                                text=[f"{v:,.0f}" if v>0 else "" for v in cvc],textposition='auto',textfont=dict(size=11),
                                hovertemplate="%{x}월: %{y:,.0f}<extra></extra>"))
                            st.plotly_chart(fm2, key=f"fm2_{usage_name}_{ks}_{sel}", width="stretch")

        render_usage("산업용","2",key_sfx)
        st.markdown("<hr style='margin:50px 0;border-top:2px solid #ccc;'>",unsafe_allow_html=True)
        render_usage("업무용","3",key_sfx)

        # ── 4. 보고서 출력 ──
        st.markdown("<hr style='border-top:2px solid #bbb;margin:40px 0 20px 0;'>",unsafe_allow_html=True)
        st.markdown("### 🖨️ 4. 보고서 출력")
        st.markdown("""<style>@media print{
            header[data-testid="stHeader"]{display:none!important;}
            section[data-testid="stSidebar"]{display:none!important;}
            div[data-testid="stToolbar"]{display:none!important;}
        }</style>""", unsafe_allow_html=True)
        st.html("""<button onclick="window.parent.print()" style="padding:12px 20px;font-size:16px;
            border-radius:8px;background-color:#1e3a8a;color:white;border:none;cursor:pointer;
            width:100%;font-weight:bold;box-shadow:0 4px 6px rgba(0,0,0,0.1);margin:2px;">
            🖨️ 현재 화면 전체를 PDF로 다운로드 (인쇄)</button>""")
