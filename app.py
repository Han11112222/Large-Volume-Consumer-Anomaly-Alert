import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="빌링 납기별 분석", layout="wide")
st.title("📊 산업용 빌링 납기 부문별 비교 분석 대시보드")

@st.cache_data
def load_and_process_data():
    encodings = ['cp949', 'utf-8', 'euc-kr', 'utf-8-sig']
    
    # 1. 메인 원장 데이터 로드 (가정용외 전체 판매량)
    main_file = "가정용외_202605.csv"
    df_main = None
    for enc in encodings:
        if os.path.exists(main_file):
            try:
                df_main = pd.read_csv(main_file, encoding=enc)
                break
            except:
                continue
    
    if df_main is None:
        st.error(f"❌ '{main_file}' 파일을 찾을 수 없습니다. 깃허브 업로드 상태를 확인해주세요.")
        return None, None, 0

    # 상품명에서 '산업용' 포함된 데이터만 정렬 및 필터링
    df_ind_main = df_main[df_main['상품명'].str.contains('산업용', na=False)].copy()
    
    # [데이터 전처리] 사용량 텍스트 내 쉼표 제거 후 숫자형 변환
    if df_ind_main['사용량(m3)'].dtype == object:
        df_ind_main['사용량(m3)'] = df_ind_main['사용량(m3)'].astype(str).str.replace(',', '')
    df_ind_main['사용량(m3)'] = pd.to_numeric(df_ind_main['사용량(m3)'], errors='coerce').fillna(0)
    
    # 기준이 되는 전체 산업용 판매량 총합
    total_industrial_volume = df_ind_main['사용량(m3)'].sum()

    # 2. 깃허브에 올라온 분할 시트 데이터 병합
    file_list = [
        "5월 관리납기 산업용 상품..xlsx - 5월 월말.csv",
        "5월 관리납기 산업용 상품..xlsx - 5월 산업용 월말2.csv",
        "5월 관리납기 산업용 상품..xlsx - 5월 산업용2회 기타2.csv",
        "5월 관리납기 산업용 상품..xlsx - 6월 산업용2회 기타1.csv"
    ]
    
    df_sub_list = []
    for file in file_list:
        if os.path.exists(file):
            for enc in encodings:
                try:
                    tmp = pd.read_csv(file, encoding=enc)
                    df_sub_list.append(tmp)
                    break
                except:
                    continue
                    
    if not df_sub_list:
        st.error("❌ 상세 검침 데이터(.csv)들을 로드하지 못했습니다.")
        return df_ind_main, None, total_industrial_volume
        
    df_sub = pd.concat(df_sub_list, ignore_index=True)
    
    # 상세 데이터 사용량 숫자형 변환
    if df_sub['사용량'].dtype == object:
        df_sub['사용량'] = df_sub['사용량'].astype(str).str.replace(',', '')
    df_sub['사용량'] = pd.to_numeric(df_sub['사용량'], errors='coerce').fillna(0)
    
    # [검침 부문 표준화 매핑]
    def map_category(x):
        x_str = str(x)
        if '1회' in x_str: return '산업용1회'
        elif '월말' in x_str: return '산업용월말'
        elif '기타' in x_str: return '산업용기타'
        return '기타/미분류'
        
    df_sub['검침부문'] = df_sub['납기구분'].apply(map_category)
    
    return df_ind_main, df_sub, total_industrial_volume

# 데이터 처리 실행
df_main, df_sub, total_ind_volume = load_and_process_data()

if df_main is not None and df_sub is not None:
    st.success("✅ 분석용 데이터가 정상적으로 매칭되었습니다.")
    
    # ----------------------------------------------------
    # 1️⃣ 섹션: 전체 양 및 부문별 양, 비율 분석
    # ----------------------------------------------------
    st.markdown("### 📊 1. 산업용 판매량 부문별 비율 분석")
    
    # 부문별 사용량 집계
    summary_df = df_sub.groupby('검침부문')['사용량'].sum().reset_index()
    summary_df.columns = ['검침 부문', '부문별 사용량(m3)']
    
    # 원장 전체 산업용 양 대비 차지하는 비율 산출
    summary_df['전체 대비 비율(%)'] = (summary_df['부문별 사용량(m3)'] / total_ind_volume) * 100
    
    # 대시보드 상단 요약 지표 (KPI Metrics)
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("원장 기준 산업용 전체 판매량", f"{total_ind_volume:,.0f} m³")
    kpi2.metric("세부 검침부문 합산 사용량", f"{df_sub['사용량'].sum():,.0f} m³")
    kpi3.metric("전체 데이터 검침 매칭률", f"{(df_sub['사용량'].sum() / total_ind_volume)*100:.2f} %")
    
    col_tbl, col_cht = st.columns([5, 5])
    
    with col_tbl:
        st.markdown("##### 📋 부문별 판매량 및 점유율 요약 표")
        # 가독성을 위한 포맷팅 
        fmt_summary = summary_df.copy()
        fmt_summary['부문별 사용량(m3)'] = fmt_summary['부문별 사용량(m3)'].map('{:,.0f}'.format)
        fmt_summary['전체 대비 비율(%)'] = fmt_summary['전체 대비 비율(%)'].map('{:.2f}%'.format)
        st.dataframe(fmt_summary, use_container_width=True, hide_index=True)
        
    with col_cht:
        fig_pie = px.pie(
            summary_df,
            names='검침 부문',
            values='부문별 사용량(m3)',
            hole=0.4,
            title="전체 산업용 판매량 내 부문별 비율",
            color_discrete_sequence=px.colors.sequential.Blues_r
        )
        st.plotly_chart(fig_pie, use_container_width=True)
        
    st.divider()
    
    # ----------------------------------------------------
    # 2️⃣ 섹션: 활성화 버튼을 이용한 부문별 상위 업체 정렬
    # ----------------------------------------------------
    st.markdown("### 🏢 2. 부문별 최다 사용 업체 정렬 (활성화)")
    
    # 활성화 수단으로 직관적인 라디오 버튼 활용
    selected_sector = st.radio(
        "분석할 부문 버튼을 클릭하세요:",
        options=['산업용1회', '산업용월말', '산업용기타'],
        horizontal=True
    )
    
    if selected_sector:
        # 선택된 부문 데이터 필터링
        sector_data = df_sub[df_sub['검침부문'] == selected_sector]
        
        # 고객명별 집계 후 내림차순 정렬 (Top 10)
        top_customers = sector_data.groupby('고객명')['사용량'].sum().reset_index()
        top_customers = top_customers.sort_values(by='사용량', ascending=False).head(10)
        top_customers.columns = ['고객명', '사용량(m3)']
        
        st.markdown(f"#### 🔍 **{selected_sector}** 부문 사용량 상위 10개 업체 현황")
        
        col_bar, col_list = st.columns([6, 4])
        
        with col_bar:
            # 시각적인 가로 바 차트
            fig_bar = px.bar(
                top_customers,
                x='사용량(m3)',
                y='고객명',
                orientation='h',
                text='사용량(m3)',
                color='사용량(m3)',
                color_continuous_scale=px.colors.sequential.GnBu,
                title=f"{selected_sector} Top 10 업체 순위"
            )
            fig_bar.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)
            
        with col_list:
            # 깔끔하게 가공된 랭킹 리스트 표
            fmt_top = top_customers.copy().reset_index(drop=True)
            fmt_top.index = fmt_top.index + 1 # 순위를 1부터 표시
            fmt_top['사용량(m3)'] = fmt_top['사용량(m3)'].map('{:,.0f}'.format)
            st.markdown("##### 🏆 사용량 순위표")
            st.dataframe(fmt_top, use_container_width=True)
