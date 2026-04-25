import io
import random
from pathlib import Path
from typing import Tuple

import numpy as np
import pandas as pd
import matplotlib as mpl
import plotly.graph_objects as go
import pydeck as pdk
import requests
import streamlit as st

# 🟢 [수정] 최상단 위치 고정
st.set_page_config(page_title="대용량 수요처 이상 감지 대시보드", layout="wide")

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

# ─────────────────────────────────────────────────────────
# 데이터 전처리 유틸
# ─────────────────────────────────────────────────────────
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

def load_safe_csv(file_bytes) -> pd.DataFrame:
    encodings = ["utf-8-sig", "cp949", "utf-8", "euc-kr"]
    for enc in encodings:
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, thousands=',')
            df.columns = df.columns.astype(str).str.strip().str.replace('\ufeff', '')
            return df
        except Exception:
            pass
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def geocode_address(address: str, api_key: str = "") -> Tuple[float, float]:
    if pd.isna(address) or not str(address).strip():
        return None, None
    if api_key:
        url = f"https://dapi.kakao.com/v2/local/search/address.json?query={address}"
        headers = {"Authorization": f"KakaoAK {api_key}"}
        try:
            res = requests.get(url, headers=headers).json()
            if res.get('documents'):
                return float(res['documents'][0]['y']), float(res['documents'][0]['x'])
        except Exception:
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
    up_csvs = st.file_uploader("가정용외_*.csv 형식 업로드", type=["csv"], accept_multiple_files=True)
    if up_csvs:
        df_list = []
        for f in up_csvs:
            df = load_safe_csv(f.getvalue())
            if not df.empty: df_list.append(df)
        if df_list: st.session_state['merged_csv_df'] = pd.concat(df_list, ignore_index=True)
    
    st.markdown("---")
    kakao_key = st.text_input("카카오 REST API 키", type="password")

# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────
if 'merged_csv_df' in st.session_state:
    df_csv = st.session_state['merged_csv_df'].copy()
    
    if not df_csv.empty:
        if "사용량(mj)" in df_csv.columns: df_csv["사용량(mj)"] = df_csv["사용량(mj)"].apply(clean_korean_finance_number)
        if "사용량(m3)" in df_csv.columns: df_csv["사용량(m3)"] = df_csv["사용량(m3)"].apply(clean_korean_finance_number)
            
    rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

    for idx, rpt_tab in enumerate(rpt_tabs):
        with rpt_tab:
            unit_str = "GJ" if idx == 0 else "천m³"
            val_col = "사용량(mj)" if idx == 0 else "사용량(m3)"
            key_sfx = f"_{idx}"

            st.markdown(f"#### 📅 기준 일자 설정") 
            
            df_csv_tab = df_csv.copy()
            if not df_csv_tab.empty:
                # 🟢 GJ 단위 변환 고정
                if unit_str == "GJ":
                    df_csv_tab[val_col] = df_csv_tab[val_col] / 1000.0
                else:
                    df_csv_tab[val_col] = df_csv_tab[val_col] / 1000.0
                
                # 날짜 처리
                for c in ["청구년월", "매출년월", "년월"]:
                    if c in df_csv_tab.columns:
                        df_csv_tab["date_parsed"] = pd.to_datetime(df_csv_tab[c], errors='coerce')
                        break
                df_csv_tab["연_csv"] = df_csv_tab["date_parsed"].dt.year
                df_csv_tab["월_csv"] = df_csv_tab["date_parsed"].dt.month
                
                avail_years = sorted(df_csv_tab["연_csv"].dropna().unique().astype(int).tolist())
                sel_year = st.selectbox("기준 연도", avail_years, index=len(avail_years)-1, key=f"yr{key_sfx}")
                sel_month = st.selectbox("기준 월", list(range(1, 13)), index=2, key=f"mo{key_sfx}")

                st.markdown("---")

                # 🟢 [수정 포인트 1] 순서 변경: 산업용(1) -> 업무용(2)
                # 🟢 [수정 포인트 2] 코멘트 섹션 삭제됨
                usages = ["산업용", "업무용"]
                for u_idx, usage_name in enumerate(usages):
                    st.subheader(f"{u_idx+1}. 용도별 판매량 분석 : {usage_name}")
                    
                    if usage_name == "업무용":
                        df_u = df_csv_tab[(df_csv_tab["용도"] == usage_name) | (df_csv_tab["상품명"].astype(str).str.contains("업무|미군", na=False))]
                    else:
                        df_u = df_csv_tab[df_csv_tab["용도"] == usage_name]
                    
                    df_u_filtered = df_u[df_u["월_csv"] <= sel_month]
                    
                    if not df_u_filtered.empty:
                        curr_vals = df_u_filtered[df_u_filtered["연_csv"] == sel_year].groupby("월_csv")[val_col].sum()
                        prev_vals = df_u_filtered[df_u_filtered["연_csv"] == sel_year-1].groupby("월_csv")[val_col].sum()
                        
                        # 차트 레이아웃
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            fig_cum = go.Figure(data=[go.Bar(x=["당해", "전년"], y=[curr_vals.sum(), prev_vals.sum()], marker_color=[COLOR_ACT, COLOR_PREV], text=[f"{curr_vals.sum():,.0f}", f"{prev_vals.sum():,.0f}"], textposition='auto')])
                            fig_cum.update_layout(height=350, title=f"누적 비교 ({unit_str})")
                            st.plotly_chart(fig_cum, use_container_width=True, key=f"cum_{usage_name}{key_sfx}")
                        with c2:
                            fig_mo = go.Figure()
                            fig_mo.add_trace(go.Bar(x=curr_vals.index, y=curr_vals.values, name="당해", marker_color=COLOR_ACT))
                            fig_mo.add_trace(go.Bar(x=prev_vals.index, y=prev_vals.values, name="전년", marker_color=COLOR_PREV))
                            fig_mo.update_layout(barmode='group', height=350, title=f"월별 추이 ({unit_str})")
                            st.plotly_chart(fig_mo, use_container_width=True, key=f"mo_{usage_name}{key_sfx}")

                        # 🟢 [수정 포인트 3] 그래프에서 '기타' 항목 제외
                        grp_col = "업종" if "업종" in df_u.columns else "업종분류"
                        u_curr = df_u_filtered[df_u_filtered["연_csv"] == sel_year].groupby(grp_col)[val_col].sum()
                        u_prev = df_u_filtered[df_u_filtered["연_csv"] == sel_year-1].groupby(grp_col)[val_col].sum()
                        u_comp = pd.merge(u_curr, u_prev, on=grp_col, how='outer', suffixes=('_c', '_p')).fillna(0)
                        u_comp = u_comp.sort_values('_c', ascending=False).head(10) # '기타' 없이 순수 Top 10만!
                        
                        fig_grp = go.Figure()
                        fig_grp.add_trace(go.Bar(x=u_comp.index, y=u_comp['_c'], name="당해", marker_color=COLOR_ACT))
                        fig_grp.add_trace(go.Bar(x=u_comp.index, y=u_comp['_p'], name="전년", marker_color=COLOR_PREV))
                        fig_grp.update_layout(barmode='group', height=400, title=f"{usage_name} 업종별 비교 (Top 10)")
                        st.plotly_chart(fig_grp, use_container_width=True, key=f"grp_{usage_name}{key_sfx}")
                        
                        # 별첨 섹션 (원본 구조 유지)
                        with st.expander(f"🔍 {usage_name} 세부 분석 및 고객 보기"):
                            c_curr = df_u_filtered[df_u_filtered["연_csv"] == sel_year].groupby("고객명")[val_col].sum()
                            c_prev = df_u_filtered[df_u_filtered["연_csv"] == sel_year-1].groupby("고객명")[val_col].sum()
                            c_comp = pd.merge(c_curr, c_prev, on="고객명", how='outer', suffixes=('_c', '_p')).fillna(0)
                            c_comp["증감"] = c_comp["_c"] - c_comp["_p"]
                            c_comp = c_comp.sort_values("_c", ascending=False).head(30)
                            st.dataframe(center_style(c_comp.style.format("{:,.0f}")), use_container_width=True)

                st.markdown("---")

            # 3. 지도 시각화
            st.subheader("3. 대용량 수요처 이상 감지 모니터링 지도")
            map_data = df_csv_tab[(df_csv_tab["연_csv"] == sel_year) & (df_csv_tab["월_csv"] == sel_month)].copy()
            if not map_data.empty and "도로명주소" in map_data.columns:
                map_data['lat_lon'] = map_data["도로명주소"].apply(lambda x: geocode_address(x, kakao_key))
                map_data = map_data.dropna(subset=['lat_lon'])
                map_data['lat'] = map_data['lat_lon'].apply(lambda x: x[0])
                map_data['lon'] = map_data['lat_lon'].apply(lambda x: x[1])
                
                layer = pdk.Layer("ScatterplotLayer", map_data, get_position='[lon, lat]', get_color=COLOR_ALARM, get_radius=200, pickable=True)
                st.pydeck_chart(pdk.Deck(layers=[layer], initial_view_state=pdk.ViewState(latitude=35.87, longitude=128.6, zoom=11, pitch=45)))
else:
    st.info("👈 사이드바에서 CSV 파일을 업로드해 주세요.")
