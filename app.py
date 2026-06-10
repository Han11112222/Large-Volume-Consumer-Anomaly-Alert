import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="빌링 납기 분석", layout="wide")
st.title("📊 산업용 판매량 부문별 종합 분석")

@st.cache_data
def load_and_process_data():
    encodings = ['cp949', 'utf-8', 'euc-kr', 'utf-8-sig']
    
    # 1. 메인 원장 데이터 (전체 판매량의 기준)
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
        st.error("❌ '가정용외_202605.csv' 파일을 찾을 수 없습니다.")
        return None, None, 0

    # '상품명'에서 '산업용'만 필터링하여 진짜 "전체 산업용 판매량" 계산
    df_ind_main = df_main[df_main['상품명'].str.contains('산업용', na=False)].copy()
    if df_ind_main['사용량(m3)'].dtype == object:
        df_ind_main['사용량(m3)'] = df_ind_main['사용량(m3)'].astype(str).str.replace(',', '')
    df_ind_main['사용량(m3)'] = pd.to_numeric(df_ind_main['사용량(m3)'], errors='coerce').fillna(0)
    
    total_industrial_volume = df_ind_main['사용량(m3)'].sum()

    # 2. 세부 검침 부문 데이터 병합
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
        st.error("❌ 세부 검침 데이터(.csv)들을 찾을 수 없습니다.")
        return df_ind_main, None, total_industrial_volume
        
    df_sub = pd.concat(df_sub_list, ignore_index=True)
    
    # 사용량 숫자형 변환
    if df_sub['사용량'].dtype == object:
        df_sub['사용량'] = df_sub['사용량'].astype(str).str.replace(',', '')
    df_sub['사용량'] = pd.to_numeric(df_sub['사용량'], errors='coerce').fillna(0)
    
    # 검침 부문 매핑
    def map_category(x):
        x_str = str(x)
        if '1회' in x_str: return '산업용1회'
        elif '월말' in x_str: return '산업용월말'
        elif '기타' in x_str: return '산업용기타'
        return '기타'
        
    df_sub['검침부문'] = df_sub['납기구분'].apply(map_category)
    
    return df_ind_main, df_sub, total_industrial_volume

# 실행
df_main, df_sub, total_volume = load_and_process_data()

if df_main is not None and df_sub is not None:
    
    # --- 1. 부문별 합계 및 비율 계산 ---
    summary_df = df_sub.groupby('검침부문')['사용량'].sum().reset_index()
    summary_df.columns = ['검침부문', '부문별 판매량(m3)']
    
    # 세부 데이터 합산량이 전체 원장과 차이가 날 경우 '미매칭 물량'으로 표기
    sub_total = summary_df['부문별 판매량(m3)'].sum()
    diff_volume = total_volume - sub_total
    
    if diff_volume > 1: # 약간의 소수점 오차 제외
        unmatched_df = pd.DataFrame({
            '검침부문': ['미매칭/기타 물량'],
            '부문별 판매량(m3)': [diff_volume]
        })
        summary_df = pd.concat([summary_df, unmatched_df], ignore_index=True)
        
    # 전체 대비 비율 계산
    summary_df['전체대비 비율(%)'] = (summary_df['부문별 판매량(m3)'] / total_volume) * 100

    # --- 2. 화면 출력 ---
    st.markdown(f"### 🎯 기준: 산업용 전체 판매량 총합 **{total_volume:,.0f} m³**")
    st.info("가정용외 원장(202605) 파일의 산업용 전체 판매량을 기준으로, 각 검침 부문이 차지하는 실제 양(m³)과 비율(%)을 분석합니다.")
    
    col_table, col_chart = st.columns([4, 6])
    
    with col_table:
        st.markdown("#### 📋 부문별 판매량 요약")
        # 가독성을 위한 포맷 적용
        display_df = summary_df.copy()
        display_df['부문별 판매량(m3)'] = display_df['부문별 판매량(m3)'].map('{:,.0f}'.format)
        display_df['전체대비 비율(%)'] = display_df['전체대비 비율(%)'].map('{:.2f}%'.format)
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    with col_chart:
        # Plotly 도넛 차트 (양과 비율을 동시에 표시)
        fig = px.pie(
            summary_df,
            names='검침부문',
            values='부문별 판매량(m3)',
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Teal
        )
        # 텍스트 템플릿을 수정하여 이름, 양(m3), 퍼센트 모두 직관적으로 보이게 설정
        fig.update_traces(
            textinfo='label+percent+value',
            texttemplate='<b>%{label}</b><br>%{value:,.0f} m³<br>(%{percent})'
        )
        fig.update_layout(showlegend=False, margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    
    # --- 3. 하단: 활성화 버튼(라디오)을 통한 상위 업체 조회 ---
    st.markdown("### 🏢 부문별 최다 사용 업체 조회")
    
    selected_sector = st.radio(
        "상세히 볼 부문을 선택하세요:",
        options=['산업용1회', '산업용월말', '산업용기타'],
        horizontal=True
    )
    
    if selected_sector:
        sector_data = df_sub[df_sub['검침부문'] == selected_sector]
        top_customers = sector_data.groupby('고객명')['사용량'].sum().reset_index()
        top_customers = top_customers.sort_values(by='사용량', ascending=False).head(10)
        
        fig_bar = px.bar(
            top_customers,
            x='사용량',
            y='고객명',
            orientation='h',
            text='사용량',
            title=f"[{selected_sector}] 상위 10개 업체 사용량 (m³)",
            color='사용량',
            color_continuous_scale=px.colors.sequential.Teal
        )
        fig_bar.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
        fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)
