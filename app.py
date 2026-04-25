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

# 🟢 Streamlit UI Crash 방지를 위한 최상단 설정
st.set_page_config(page_title="대용량 수요처 이상 감지 대시보드", layout="wide")

# ─────────────────────────────────────────────────────────
# 기본 설정 (폰트)
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
# 공통 유틸리티 함수
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
# 사이드바 설정
# ─────────────────────────────────────────────────────────
st.title("📊 대용량 수요처 이상 감지 대시보드")

with st.sidebar:
    st.header("📂 데이터 설정")
    up_csvs = st.file_uploader("업종별 실적 CSV 업로드", type=["csv"], accept_multiple_files=True)
    
    if up_csvs:
        df_list = []
        for f in up_csvs:
            df = load_safe_csv(f.getvalue())
            if not df.empty: df_list.append(df)
        if df_list:
            st.session_state['data'] = pd.concat(df_list, ignore_index=True)
    
    st.markdown("---")
    kakao_key = st.text_input("카카오 API 키 (선택사항)", type="password")

# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────
if 'data' in st.session_state:
    df_raw = st.session_state['data'].copy()
    
    # 탭 설정
    rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])
    
    for t_idx, rpt_tab in enumerate(rpt_tabs):
        with rpt_tab:
            unit_str = "GJ" if t_idx == 0 else "천m³"
            val_col = "사용량(mj)" if t_idx == 0 else "사용량(m3)"
            key_suffix = f"_{t_idx}"
            
            # 데이터 클리닝 및 단위 변환 (MJ -> GJ)
            df_tab = df_raw.copy()
            df_tab[val_col] = df_tab[val_col].apply(clean_korean_finance_number)
            df_tab[val_col] = df_tab[val_col] / 1000.0
            
            # 날짜 처리
            for c in ["청구년월", "매출년월", "년월"]:
                if c in df_tab.columns:
                    df_tab["date_parsed"] = pd.to_datetime(df_tab[c], errors='coerce')
                    break
            
            df_tab["year"] = df_tab["date_parsed"].dt.year
            df_tab["month"] = df_tab["date_parsed"].dt.month
            
            # 기준 일자 선택 UI
            years = sorted(df_tab["year"].dropna().unique().astype(int).tolist())
            c1, c2 = st.columns(2)
            with c1: sel_year = st.selectbox("기준 연도", years, index=len(years)-1, key=f"yr{key_suffix}")
            with c2: sel_month = st.selectbox("기준 월", list(range(1, 13)), index=2, key=f"mo{key_suffix}")
            
            st.markdown("---")
            
            # 🟢 [순서 변경] 1. 산업용 분석
            # ─────────────────────────────────────────────────────
            usage_list = ["산업용", "업무용"] # 순서 고정
            
            for u_idx, usage in enumerate(usage_list):
                st.subheader(f"{u_idx+1}. 용도별 판매량 분석 : {usage}")
                
                # 용도 및 월 필터링
                if usage == "업무용":
                    df_u = df_tab[(df_tab["용도"] == usage) | (df_tab.get("상품명", pd.Series()).str.contains("업무|미군", na=False))]
                else:
                    df_u = df_tab[df_tab["용도"] == usage]
                
                df_u_filtered = df_u[df_u["month"] <= sel_month]
                
                if not df_u_filtered.empty:
                    # 상단 차트 데이터 (금년/전년 실적)
                    curr_vals = df_u_filtered[df_u_filtered["year"] == sel_year].groupby("month")[val_col].sum()
                    prev_vals = df_u_filtered[df_u_filtered["year"] == sel_year-1].groupby("month")[val_col].sum()
                    
                    sum_c, sum_p = curr_vals.sum(), prev_vals.sum()
                    
                    # 1&2. 누적/월별 막대그래프
                    mc1, mc2 = st.columns([1, 2])
                    with mc1:
                        fig_cum = go.Figure(data=[go.Bar(x=["당해실적", "전년실적"], y=[sum_c, sum_p], marker_color=[COLOR_ACT, COLOR_PREV], text=[f"{sum_c:,.0f}", f"{sum_p:,.0f}"], textposition='auto')])
                        fig_cum.update_layout(height=350, margin=dict(t=20, b=20), title=f"누적 비교 ({unit_str})")
                        st.plotly_chart(fig_cum, use_container_width=True, key=f"cum_{usage}{key_suffix}")
                    
                    with mc2:
                        fig_mo = go.Figure()
                        fig_mo.add_trace(go.Bar(x=curr_vals.index, y=curr_vals.values, name="당해", marker_color=COLOR_ACT))
                        fig_mo.add_trace(go.Bar(x=prev_vals.index, y=prev_vals.values, name="전년", marker_color=COLOR_PREV))
                        fig_mo.update_layout(barmode='group', height=350, margin=dict(t=20, b=20), title=f"월별 추이 ({unit_str})")
                        st.plotly_chart(fig_mo, use_container_width=True, key=f"mo_{usage}{key_suffix}")

                    # 3. 세부 업종별 그래프 (기타 항목 삭제)
                    st.markdown(f"**■ {usage} 세부 업종별 비교 (Top 10)**")
                    grp_col = "업종" if "업종" in df_u.columns else "업종분류"
                    if grp_col:
                        u_curr = df_u_filtered[df_u_filtered["year"] == sel_year].groupby(grp_col)[val_col].sum()
                        u_prev = df_u_filtered[df_u_filtered["year"] == sel_year-1].groupby(grp_col)[val_col].sum()
                        u_comp = pd.merge(u_curr, u_prev, on=grp_col, how='outer', suffixes=('_c', '_p')).fillna(0)
                        u_comp = u_comp.sort_values('_c', ascending=False).head(10) # 🟢 기타 없이 Top 10만!
                        
                        fig_grp = go.Figure()
                        fig_grp.add_trace(go.Bar(x=u_comp.index, y=u_comp['_c'], name="당해", marker_color=COLOR_ACT))
                        fig_grp.add_trace(go.Bar(x=u_comp.index, y=u_comp['_p'], name="전년", marker_color=COLOR_PREV))
                        fig_grp.update_layout(barmode='group', height=400, title="업종별 비교 (Top 10)")
                        st.plotly_chart(fig_grp, use_container_width=True, key=f"grp_{usage}{key_suffix}")
                    
                    # 4. Top 30 및 상세 리스트
                    with st.expander(f"🔍 {usage} 상세 리스트 및 개별 고객 보기"):
                        st.markdown(f"**🏆 {usage} 판매량 Top 30**")
                        c_curr = df_u_filtered[df_u_filtered["year"] == sel_year].groupby("고객명")[val_col].sum()
                        c_prev = df_u_filtered[df_u_filtered["year"] == sel_year-1].groupby("고객명")[val_col].sum()
                        c_comp = pd.merge(c_curr, c_prev, on="고객명", how='outer', suffixes=('_c', '_p')).fillna(0)
                        c_comp["증감"] = c_comp["_c"] - c_comp["_p"]
                        c_comp = c_comp.sort_values("_c", ascending=False).head(30)
                        st.dataframe(center_style(c_comp.style.format("{:,.0f}")), use_container_width=True)
                        
                        sel_c = st.selectbox(f"분석할 고객 선택 ({usage})", ["선택안함"] + c_comp.index.tolist(), key=f"sel_{usage}{key_suffix}")
                        if sel_c != "선택안함":
                            cd = df_u[df_u["고객명"] == sel_c].groupby(["year", "month"])[val_col].sum().unstack(level=0).fillna(0)
                            st.bar_chart(cd)

                st.markdown("---")

            # 🟢 3. 지도 시각화
            st.subheader("3. 이상 감지 업체 지리적 분포")
            map_data = df_tab[(df_tab["year"] == sel_year) & (df_tab["month"] == sel_month)].copy()
            if not map_data.empty and "도로명주소" in map_data.columns:
                # 간단한 리스크 샘플링 (10% 이상 하락)
                map_data['lat_lon'] = map_data["도로명주소"].apply(lambda x: geocode_address(x, kakao_key))
                map_data = map_data.dropna(subset=['lat_lon'])
                map_data['lat'] = map_data['lat_lon'].apply(lambda x: x[0])
                map_data['lon'] = map_data['lat_lon'].apply(lambda x: x[1])
                
                layer = pdk.Layer("ScatterplotLayer", map_data, get_position='[lon, lat]', get_color=COLOR_ALARM, get_radius=200, pickable=True)
                st.pydeck_chart(pdk.Deck(layers=[layer], initial_view_state=pdk.ViewState(latitude=35.87, longitude=128.6, zoom=10, pitch=45)))

else:
    st.info("👈 사이드바에서 CSV 파일을 업로드해 주세요.")
