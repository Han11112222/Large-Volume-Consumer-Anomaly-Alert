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

# 🟢 [중요] 화면 하얗게 뻗는 버그 방지: 무조건 제일 위로 올렸습니다.
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

# 🟢 KeyError 완벽 차단 로직
def get_usage_data(df, usage_name):
    if df is None or df.empty or "용도" not in df.columns:
        return pd.DataFrame()
    
    if usage_name == "산업용":
        return df[df["용도"] == "산업용"].copy()
    elif usage_name == "업무용":
        if "상품명" in df.columns:
            prod_series = df["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
            mask = (df["용도"] == "업무용") | (prod_series.isin(["냉난방용(업무)", "업무난방용", "주한미군"]))
            return df[mask].copy()
        else:
            return df[df["용도"] == "업무용"].copy()
    else:
        return df[df["용도"] == usage_name].copy()

# ─────────────────────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────────────────────
st.title("📊 대용량 수요처 이상 감지 대시보드")

with st.sidebar:
    st.header("📂 데이터 및 설정")

    st.subheader("1. 업종별 데이터 (필수/CSV)")
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
    st.subheader("🗺️ 지도 API 키")
    kakao_key = st.text_input("카카오 REST API 키", type="password", help="키가 없으면 대구 임의의 좌표로 매핑됩니다.")

# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────
try:
    df_csv = pd.DataFrame()

    if src_csv == "레포 파일 사용":
        repo_dir = Path(__file__).parent
        all_csvs = list(repo_dir.glob("*가정용외*.csv")) + list(repo_dir.glob("가정용외*.csv"))
        all_csvs = list(set(all_csvs)) 
        csv_list = []
        for p in all_csvs:
            try:
                with open(p, 'rb') as f:
                    temp_df = load_safe_csv(f.read())
                    if not temp_df.empty:
                        csv_list.append(temp_df)
            except:
                pass
        if csv_list: df_csv = pd.concat(csv_list, ignore_index=True)

    if df_csv.empty and 'merged_csv_df' in st.session_state:
        df_csv = st.session_state['merged_csv_df'].copy()
        
    if not df_csv.empty:
        if "사용량(mj)" in df_csv.columns: df_csv["사용량(mj)"] = df_csv["사용량(mj)"].apply(clean_korean_finance_number)
        if "사용량(m3)" in df_csv.columns: df_csv["사용량(m3)"] = df_csv["사용량(m3)"].apply(clean_korean_finance_number)
            
    rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

    for idx, rpt_tab in enumerate(rpt_tabs):
        with rpt_tab:
            if idx == 0:
                unit_str = "GJ"
                val_col = "사용량(mj)"
                key_sfx = "_gj"
            else:
                unit_str = "천m³"
                val_col = "사용량(m3)"
                key_sfx = "_vol"

            st.markdown(f"#### 📅 기준 일자 설정") 
            
            years_available = [2024, 2025, 2026]
            default_y_index = len(years_available) - 1
            default_m_index = 2 
            
            df_csv_tab = df_csv.copy()
            if not df_csv_tab.empty:
                # 🟢 GJ 변환
                if unit_str == "GJ" and "사용량(mj)" in df_csv_tab.columns:
                    df_csv_tab["사용량(mj)"] = pd.to_numeric(df_csv_tab["사용량(mj)"], errors="coerce").fillna(0) / 1000.0
                elif unit_str == "천m³" and "사용량(m3)" in df_csv_tab.columns:
                    df_csv_tab["사용량(m3)"] = pd.to_numeric(df_csv_tab["사용량(m3)"], errors="coerce").fillna(0) / 1000.0
                    
                # 날짜 파싱
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
                
                # 🟢 TypeError 원천 차단: 어떤 데이터 형태가 오더라도 오로지 tolist()로 완전한 리스트만 추출
                try:
                    if "연_csv" in df_csv_tab.columns:
                        valid_years = df_csv_tab["연_csv"].dropna().astype(int)
                        if not valid_years.empty:
                            years_available = sorted(list(set(valid_years.tolist())))
                except Exception:
                    years_available = [2024, 2025, 2026]

                if years_available:
                    max_year = max(years_available)
                    
                    try:
                        max_month = int(df_csv_tab[df_csv_tab["연_csv"] == max_year]["월_csv"].max())
                        if pd.isna(max_month): max_month = 3
                    except:
                        max_month = 3
                        
                    default_y_index = years_available.index(max_year) if max_year in years_available else len(years_available) - 1
                    default_m_index = max(0, max_month - 1)
            
            c_y, c_m, c_empty = st.columns([1, 1, 2])
            with c_y:
                # 🟢 years_available 리스트가 혹시나 비었을 경우를 방어
                if not years_available:
                    years_available = [2026]
                    default_y_index = 0
                sel_year_rpt = st.selectbox("기준 연도", years_available, index=default_y_index, key=f"rpt_yr{key_sfx}")
            with c_m:
                sel_month_str = st.selectbox("기준 월", [f"{m}월" for m in range(1, 13)], index=default_m_index, key=f"rpt_mo{key_sfx}")
            
            max_month = int(sel_month_str.replace("월", "")) 
            
            st.markdown("<hr style='margin: 10px 0 30px 0;'>", unsafe_allow_html=True)

            # ─────────────────────────────────────────────────────────
            # 통합 분석 함수
            # ─────────────────────────────────────────────────────────
            def render_full_usage_report(usage_name, section_num, key_sfx):
                st.markdown(f"""<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 10px;"><h4 style="margin: 0;">📈 {section_num}. 용도별 판매량 분석 : {usage_name}</h4></div>""", unsafe_allow_html=True)
                
                if df_csv_tab.empty or "용도" not in df_csv_tab.columns or val_col not in df_csv_tab.columns:
                    st.info("데이터가 부족하여 차트를 표시할 수 없습니다.")
                    return

                df_u = get_usage_data(df_csv_tab, usage_name)
                df_u = df_u[df_u["월_csv"] <= max_month]
                
                if df_u.empty:
                    st.warning(f"업로드된 데이터에 '{usage_name}' 데이터가 없습니다.")
                    return
                
                p_curr_act = df_u[df_u["연_csv"] == sel_year_rpt].groupby("월_csv")[val_col].sum()
                p_prev_act = df_u[df_u["연_csv"] == sel_year_rpt-1].groupby("월_csv")[val_col].sum()
                
                sum_act = p_curr_act.sum()
                sum_prev = p_prev_act.sum()
                
                diff_prev = sum_act - sum_prev
                rate_prev = (sum_act / sum_prev * 100) if sum_prev > 0 else 0
                sign_prev = "+" if diff_prev > 0 else ""
                
                months_list = list(range(1, max_month + 1))
                
                # --- 1 & 2. 누적 비교 / 월별 비교 ---
                col_c, col_m = st.columns([1, 2.5])
                with col_c:
                    st.markdown(f"**■ 누적 실적 비교 ({max_month}월 누적)**")
                    st.markdown(
                        f"""
                        <div style="background-color: #e2e8f0; border-left: 5px solid #1e3a8a; padding: 10px 10px; margin-bottom: 0px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                            <div style="font-size: 14.5px; color: #1e3a8a; font-weight: 700; line-height: 1.5;">
                                당해 실적: {sum_act:,.0f} {unit_str}<br>
                                전년대비: {sign_prev}{diff_prev:,.0f} ({rate_prev:.1f}%)
                            </div>
                        </div>
                        """, unsafe_allow_html=True
                    )
                    
                    fig_c = go.Figure()
                    fig_c.add_trace(go.Bar(x=[f"{sel_year_rpt}년<br>실적", f"{sel_year_rpt-1}년<br>실적"],
                                           y=[sum_act, sum_prev],
                                           marker_color=[COLOR_ACT, COLOR_PREV],
                                           text=[f"{sum_act:,.0f}", f"{sum_prev:,.0f}"],
                                           textposition='auto', textfont=dict(size=14)))
                    fig_c.update_layout(margin=dict(t=25, b=10, l=10, r=10), height=420, showlegend=False)
                    st.plotly_chart(fig_c, use_container_width=True, key=f"fig_c_{usage_name}_{key_sfx}")
                    
                with col_m:
                    st.markdown("**■ 월별 실적 비교 (YoY)**")
                    st.markdown("<div style='padding: 1px; margin-bottom: 27px; line-height: 1.5;'>&nbsp;<br>&nbsp;</div>", unsafe_allow_html=True)
                    
                    fig_m = go.Figure()
                    vals_act = [p_curr_act.get(m, 0) for m in months_list]
                    vals_prev = [p_prev_act.get(m, 0) for m in months_list]
                    
                    fig_m.add_trace(go.Bar(x=months_list, y=vals_act, name=f'{sel_year_rpt}년 실적', marker_color=COLOR_ACT, text=[f"{v:,.0f}" if v>0 else "" for v in vals_act], textposition='auto', textfont=dict(size=11)))
                    fig_m.add_trace(go.Bar(x=months_list, y=vals_prev, name=f'{sel_year_rpt-1}년 실적', marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v>0 else "" for v in vals_prev], textposition='auto', textfont=dict(size=11)))
                    
                    fig_m.update_layout(barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1), xaxis_title="월", yaxis_title=f"판매량({unit_str})", margin=dict(t=10, b=10, l=10, r=10), height=420, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                    st.plotly_chart(fig_m, use_container_width=True, key=f"fig_m_{usage_name}_{key_sfx}")

                # --- 3. 세부 업종별 판매량 비교 (그래프) ---
                grp_col = "업종분류" if "업종분류" in df_u.columns else ("업종" if "업종" in df_u.columns else None)
                    
                if grp_col:
                    st.markdown(f"**■ 세부 업종별 판매량 비교 (당해연도 vs 전년도)**")
                    curr_ind_grp = df_u[df_u["연_csv"] == sel_year_rpt].groupby(grp_col, as_index=False)[val_col].sum().rename(columns={val_col: f"{sel_year_rpt}년"})
                    prev_ind_grp = df_u[df_u["연_csv"] == sel_year_rpt - 1].groupby(grp_col, as_index=False)[val_col].sum().rename(columns={val_col: f"{sel_year_rpt-1}년"})
                    
                    ind_comp_graph = pd.merge(curr_ind_grp, prev_ind_grp, on=grp_col, how="outer").fillna(0)
                    ind_comp_graph = ind_comp_graph.sort_values(f"{sel_year_rpt}년", ascending=False).reset_index(drop=True)
                    
                    if len(ind_comp_graph) > 10:
                        ind_comp_plot = ind_comp_graph.iloc[:10].copy()
                    else:
                        ind_comp_plot = ind_comp_graph.copy()
                            
                    ind_comp_plot["증감절대값"] = abs(ind_comp_plot[f"{sel_year_rpt}년"] - ind_comp_plot[f"{sel_year_rpt-1}년"])
                    
                    # 🟢 방어코드: 빈 데이터프레임일 때 에러 방지
                    if not ind_comp_plot.empty:
                        max_diff_idx = ind_comp_plot["증감절대값"].idxmax()
                        colors_act = [COLOR_ACT] * len(ind_comp_plot)
                        if pd.notna(max_diff_idx) and max_diff_idx is not None: 
                            colors_act[int(max_diff_idx)] = "#d32f2f" 
                    else:
                        colors_act = []
                        
                    fig_ind = go.Figure()
                    fig_ind.add_trace(go.Bar(x=ind_comp_plot[grp_col], y=ind_comp_plot[f"{sel_year_rpt}년"], name=f'{sel_year_rpt}년', marker_color=colors_act, text=[f"{v:,.0f}" if v>0 else "" for v in ind_comp_plot[f"{sel_year_rpt}년"]], textposition='auto', textfont=dict(size=11)))
                    fig_ind.add_trace(go.Bar(x=ind_comp_plot[grp_col], y=ind_comp_plot[f"{sel_year_rpt-1}년"], name=f'{sel_year_rpt-1}년', marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v>0 else "" for v in ind_comp_plot[f"{sel_year_rpt-1}년"]], textposition='auto', textfont=dict(size=11)))
                    
                    fig_ind.update_layout(barmode='group', xaxis_title="", yaxis_title=f"판매량({unit_str})", margin=dict(t=10, b=10, l=10, r=10), height=420, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                    st.plotly_chart(fig_ind, use_container_width=True, key=f"fig_ind_{usage_name}_{key_sfx}")

                st.markdown("<hr style='border-top: 1px dashed #ccc; margin: 30px 0;'>", unsafe_allow_html=True)

                # =========================================================
                # 4. 세부 업종별 비교표
                # =========================================================
                if grp_col:
                    st.markdown(f"**■ 🏢 {usage_name} 세부 업종별 비교표**")
                    
                    ind_comp = pd.merge(curr_ind_grp, prev_ind_grp, on=grp_col, how="outer").fillna(0)
                    ind_comp = ind_comp.sort_values(f"{sel_year_rpt}년", ascending=False).reset_index(drop=True)
                    
                    if len(ind_comp) > 10:
                        top10_df = ind_comp.iloc[:10].copy()
                        others_df = ind_comp.iloc[10:].copy()
                        o_c = others_df[f"{sel_year_rpt}년"].sum()
                        o_p = others_df[f"{sel_year_rpt-1}년"].sum()
                        o_diff = o_c - o_p
                        o_rate = (o_c / o_p * 100) if o_p > 0 else 0
                        
                        others_row = pd.DataFrame([{grp_col: "기타", f"{sel_year_rpt}년": o_c, f"{sel_year_rpt-1}년": o_p, "증감": o_diff, "대비(%)": o_rate}])
                        ind_comp_tbl = pd.concat([top10_df, others_row], ignore_index=True)
                    else:
                        ind_comp_tbl = ind_comp.copy()
                        ind_comp_tbl["증감"] = ind_comp_tbl[f"{sel_year_rpt}년"] - ind_comp_tbl[f"{sel_year_rpt-1}년"]
                        ind_comp_tbl["대비(%)"] = np.where(ind_comp_tbl[f"{sel_year_rpt-1}년"] > 0, (ind_comp_tbl[f"{sel_year_rpt}년"] / ind_comp_tbl[f"{sel_year_rpt-1}년"]) * 100, 0)
                    
                    sum_curr_tbl = ind_comp_tbl[f"{sel_year_rpt}년"].sum()
                    sum_prev_tbl = ind_comp_tbl[f"{sel_year_rpt-1}년"].sum()
                    sum_diff_tbl = sum_curr_tbl - sum_prev_tbl
                    sum_rate_tbl = (sum_curr_tbl / sum_prev_tbl * 100) if sum_prev_tbl > 0 else 0
                    
                    sub_ind_row = pd.DataFrame([{grp_col: "💡 총계", f"{sel_year_rpt}년": sum_curr_tbl, f"{sel_year_rpt-1}년": sum_prev_tbl, "증감": sum_diff_tbl, "대비(%)": sum_rate_tbl}])
                    ind_comp_tbl = pd.concat([ind_comp_tbl, sub_ind_row], ignore_index=True)
                    
                    st.dataframe(center_style(ind_comp_tbl.style.format({f"{sel_year_rpt}년": "{:,.0f}", f"{sel_year_rpt-1}년": "{:,.0f}", "증감": "{:,.0f}", "대비(%)": "{:,.1f}"}).apply(highlight_subtotal, axis=1)), use_container_width=True, hide_index=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # =========================================================
                    # 5. Top 30 리스트 
                    # =========================================================
                    st.markdown(f"**■ 🏆 {usage_name} Top 30 업체 List (당해연도 판매량 기준)**")
                    if "고객명" in df_u.columns:
                        c_curr_all = df_u[df_u["연_csv"] == sel_year_rpt].groupby(["고객명", grp_col], as_index=False)[val_col].sum().rename(columns={val_col: f"{sel_year_rpt}년"})
                        c_prev_all = df_u[df_u["연_csv"] == sel_year_rpt - 1].groupby(["고객명", grp_col], as_index=False)[val_col].sum().rename(columns={val_col: f"{sel_year_rpt-1}년"})
                        
                        grp_top = pd.merge(c_curr_all, c_prev_all, on=["고객명", grp_col], how="outer").fillna(0)
                        
                        grp_top = grp_top.sort_values(f"{sel_year_rpt}년", ascending=False).reset_index(drop=True)
                        grp_top = grp_top[(grp_top[f"{sel_year_rpt}년"] > 0) | (grp_top[f"{sel_year_rpt-1}년"] > 0)].reset_index(drop=True)

                        grp_top_30 = grp_top.head(30).copy()
                        grp_top_30["증감"] = grp_top_30[f"{sel_year_rpt}년"] - grp_top_30[f"{sel_year_rpt-1}년"]
                        grp_top_30["대비(%)"] = np.where(grp_top_30[f"{sel_year_rpt-1}년"] > 0, (grp_top_30[f"{sel_year_rpt}년"] / grp_top_30[f"{sel_year_rpt-1}년"]) * 100, 0)
                        
                        top30_sum_curr = grp_top_30[f"{sel_year_rpt}년"].sum()
                        top30_sum_prev = grp_top_30[f"{sel_year_rpt-1}년"].sum()
                        top30_diff = top30_sum_curr - top30_sum_prev
                        top30_rate = (top30_sum_curr / top30_sum_prev * 100) if top30_sum_prev > 0 else 0
                        top30_ratio = (top30_sum_curr / sum_curr_tbl * 100) if sum_curr_tbl > 0 else 0
                        
                        subtotal_row = pd.DataFrame([{"고객명": "💡 소계 (Top 30)", grp_col: f"전체대비 {top30_ratio:.1f}%", f"{sel_year_rpt}년": top30_sum_curr, f"{sel_year_rpt-1}년": top30_sum_prev, "증감": top30_diff, "대비(%)": top30_rate}])
                        grp_top_show = pd.concat([grp_top_30, subtotal_row], ignore_index=True)
                        
                        ranks = list(range(1, len(grp_top_30) + 1)) + ["-"]
                        grp_top_show.insert(0, "순위", ranks)
                        
                        st.dataframe(center_style(grp_top_show.style.format({f"{sel_year_rpt}년": "{:,.0f}", f"{sel_year_rpt-1}년": "{:,.0f}", "증감": "{:,.0f}", "대비(%)": "{:,.1f}"}).apply(highlight_subtotal, axis=1)), use_container_width=True, hide_index=True)
                        st.markdown("<br>", unsafe_allow_html=True)
                        
                        # =========================================================
                        # 6. 개별 고객 상세 차트
                        # =========================================================
                        st.markdown(f"**🔍 {usage_name} 개별 고객 상세 차트**")
                        top_customers = [c for c in grp_top["고객명"] if "💡" not in c]
                        sel_cust = st.selectbox(f"상세 분석할 고객명을 선택하세요 ({usage_name})", ["선택 안함"] + top_customers, key=f"sel_cust_{usage_name}_{key_sfx}")

                        if sel_cust != "선택 안함":
                            c_data = df_u[df_u["고객명"] == sel_cust]
                            c_grp = c_data.groupby(["연_csv", "월_csv"], as_index=False)[val_col].sum()
                            
                            y_cur = c_grp[(c_grp["연_csv"] == sel_year_rpt) & (c_grp["월_csv"] <= max_month)]
                            y_prev = c_grp[(c_grp["연_csv"] == sel_year_rpt - 1) & (c_grp["월_csv"] <= max_month)]
                            
                            sum_cur_c = y_cur[val_col].sum()
                            sum_prev_c = y_prev[val_col].sum()
                            diff_val = sum_cur_c - sum_prev_c
                            rate_val = (sum_cur_c / sum_prev_c * 100) if sum_prev_c > 0 else 0
                            sign_str = "+" if diff_val > 0 else ""
                            yoy_text = f"전년대비 증감: {sign_str}{diff_val:,.0f} ({rate_val:.1f}%)"
                            
                            cc1, cc2 = st.columns([1, 2])
                            with cc1:
                                fig_cust_cum = go.Figure()
                                fig_cust_cum.add_trace(go.Bar(x=[f"{sel_year_rpt}년", f"{sel_year_rpt-1}년"], y=[sum_cur_c, sum_prev_c], marker_color=[COLOR_ACT, COLOR_PREV], text=[f"{sum_cur_c:,.0f}", f"{sum_prev_c:,.0f}"], textposition='auto'))
                                fig_cust_cum.add_annotation(x=0.5, y=1.05, xref="paper", yref="paper", text=f"<b>{yoy_text}</b>", showarrow=False, font=dict(size=13, color="#d32f2f" if diff_val < 0 else "#1f77b4"), bgcolor="#f8f9fa", bordercolor="#d0d7e5", borderwidth=1, borderpad=4)
                                fig_cust_cum.update_layout(title=f"'{sel_cust}' 누적 사용량 ({max_month}월 누적)", margin=dict(t=50,b=10,l=10,r=10), height=350)
                                st.plotly_chart(fig_cust_cum, use_container_width=True, key=f"fig_cust_cum_{usage_name}_{key_sfx}")
                                
                            with cc2:
                                fig_cust_mon = go.Figure()
                                months_c = list(range(1, max_month + 1))
                                cur_vals = [y_cur[y_cur['월_csv']==m][val_col].sum() for m in months_c]
                                prev_vals = [y_prev[y_prev['월_csv']==m][val_col].sum() for m in months_c]
                                
                                fig_cust_mon.add_trace(go.Bar(x=months_c, y=cur_vals, name=f"{sel_year_rpt}년", marker_color=COLOR_ACT, text=[f"{v:,.0f}" if v>0 else "" for v in cur_vals], textposition='auto', textfont=dict(size=11)))
                                fig_cust_mon.add_trace(go.Bar(x=months_c, y=prev_vals, name=f"{sel_year_rpt-1}년", marker_color=COLOR_PREV, text=[f"{v:,.0f}" if v>0 else "" for v in prev_vals], textposition='auto', textfont=dict(size=11)))
                                
                                fig_cust_mon.update_layout(title=f"'{sel_cust}' 월별 사용량 추이", barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1), margin=dict(t=50,b=10,l=10,r=10), height=350, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                                st.plotly_chart(fig_cust_mon, use_container_width=True, key=f"fig_cust_mon_{usage_name}_{key_sfx}")

            # ─────────────────────────────────────────────────────────
            # 함수 실행 (🟢 산업용 1번, 업무용 2번)
            # ─────────────────────────────────────────────────────────
            render_full_usage_report("산업용", "1", key_sfx)
            st.markdown("<hr style='margin: 50px 0; border-top: 2px solid #ccc;'>", unsafe_allow_html=True)
            render_full_usage_report("업무용", "2", key_sfx)
            
            # ─────────────────────────────────────────────────────────
            # 3. 이상 감지 업체 지도 모니터링
            # ─────────────────────────────────────────────────────────
            st.markdown("<hr style='border-top: 2px solid #1e3a8a; margin: 50px 0 20px 0;'>", unsafe_allow_html=True)
            st.markdown("### 🗺️ 3. 대용량 수요처 이상 감지 모니터링 지도")
            st.caption("※ YoY 기준 10% 이상 사용량이 하락한 업체를 지도에 붉은색 마커로 표시하여 현장 방문을 유도합니다.")
            
            if not df_csv_tab.empty and "도로명주소" in df_csv_tab.columns and "고객명" in df_csv_tab.columns and val_col in df_csv_tab.columns and "용도" in df_csv_tab.columns:
                df_map_base = df_csv_tab[df_csv_tab["월_csv"] <= max_month].copy()
                
                if not df_map_base.empty:
                    if "상품명" in df_map_base.columns:
                        prod_s = df_map_base["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                        df_map_base["용도_태그"] = np.where(prod_s == "산업용", "[산업용]", 
                                                     np.where(prod_s.isin(["냉난방용(업무)", "업무난방용", "주한미군"]), "[업무용]", "[기타]"))
                    else:
                        df_map_base["용도_태그"] = "[분류없음]"

                    map_curr = df_map_base[df_map_base["연_csv"] == sel_year_rpt].groupby(["고객명", "도로명주소", "용도_태그"], as_index=False)[val_col].sum().rename(columns={val_col: "당해년도"})
                    map_prev = df_map_base[df_map_base["연_csv"] == sel_year_rpt - 1].groupby(["고객명", "도로명주소", "용도_태그"], as_index=False)[val_col].sum().rename(columns={val_col: "전년도"})
                    
                    if not map_curr.empty and not map_prev.empty:
                        df_map_merged = pd.merge(map_curr, map_prev, on=["고객명", "도로명주소", "용도_태그"], how="inner").fillna(0)
                        
                        df_map_merged["증감률(%)"] = np.where(df_map_merged["전년도"] > 0, ((df_map_merged["당해년도"] - df_map_merged["전년도"]) / df_map_merged["전년도"]) * 100, 0)
                        alarm_df = df_map_merged[df_map_merged["증감률(%)"] <= -10].copy()
                        
                        if alarm_df.empty:
                            st.success("✅ 선택한 기간 내 YoY 10% 이상 하락한 리스크 업체가 없습니다.")
                        else:
                            st.warning(f"🚨 총 **{len(alarm_df)}**개의 업체에서 10% 이상 하락 신호가 감지되었습니다.")
                            
                            alarm_df = alarm_df.sort_values(by="증감률(%)").head(30).reset_index(drop=True)
                            
                            lats, lons, tooltips = [], [], []
                            for _, row in alarm_df.iterrows():
                                lat, lon = geocode_address(row['도로명주소'], kakao_key)
                                lats.append(lat)
                                lons.append(lon)
                                
                                info = f"<b>{row['용도_태그']} {row['고객명']}</b><br/>"
                                info += f"전년: {row['전년도']:,.0f} / 당해: {row['당해년도']:,.0f}<br/>"
                                info += f"증감률: <span style='color:red; font-weight:bold;'>{row['증감률(%)']:.1f}%</span><br/>"
                                info += f"<span style='font-size:0.8em; color:gray;'>{row['도로명주소']}</span>"
                                tooltips.append(info)
                                
                            alarm_df['lat'] = lats
                            alarm_df['lon'] = lons
                            alarm_df['tooltip'] = tooltips
                            alarm_df = alarm_df.dropna(subset=['lat', 'lon'])
                            
                            if not alarm_df.empty:
                                layer = pdk.Layer(
                                    "ScatterplotLayer",
                                    data=alarm_df,
                                    get_position='[lon, lat]',
                                    get_color=COLOR_ALARM,
                                    get_radius=200,
                                    pickable=True,
                                    opacity=0.8,
                                    filled=True,
                                )
                                
                                view_state = pdk.ViewState(
                                    latitude=alarm_df['lat'].mean(),
                                    longitude=alarm_df['lon'].mean(),
                                    zoom=11,
                                    pitch=40,
                                )
                                
                                r = pdk.Deck(
                                    layers=[layer],
                                    initial_view_state=view_state,
                                    tooltip={"html": "{tooltip}", "style": {"backgroundColor": "white", "color": "black", "font-family": "NanumGothic"}}
                                )
                                st.pydeck_chart(r)
                            else:
                                st.error("주소 좌표 변환에 실패하여 지도를 표시할 수 없습니다.")
                    else:
                        st.info("비교할 과거 또는 당해 연도 데이터가 없습니다.")
            else:
                st.info("데이터에 '도로명주소', '고객명', '용도' 컬럼이 없거나 데이터가 부족하여 지도를 생성할 수 없습니다.")

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

except Exception as e:
    st.error(f"❌ 대시보드 렌더링 중 오류가 발생했습니다. 원인: {str(e)}")
    st.warning("데이터 파일의 형식이 맞지 않거나, 필수 컬럼명(사용량, 용도 등)이 누락되었을 수 있습니다.")
