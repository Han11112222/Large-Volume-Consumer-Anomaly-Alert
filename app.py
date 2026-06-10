with col_chart:
        # Plotly 도넛 차트 (양과 비율을 동시에 표시)
        fig = px.pie(
            summary_df,
            names='검침부문',
            values='부문별 판매량(m3)',
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Teal
        )
        # ✅ 수정된 부분: 텍스트를 밖으로 빼고(outside), 여백을 늘려 숫자가 잘리지 않게 강제 설정
        fig.update_traces(
            textposition='outside', 
            textinfo='label+value+percent',
            texttemplate='<b>%{label}</b><br>%{value:,.0f} m³<br>(%{percent})',
            textfont_size=14
        )
        fig.update_layout(
            showlegend=False, 
            margin=dict(t=50, b=50, l=50, r=50) # 차트 상하좌우 여백 확보
        )
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
            text='사용량', # 텍스트로 띄울 값 지정
            title=f"[{selected_sector}] 상위 10개 업체 사용량 (m³)",
            color='사용량',
            color_continuous_scale=px.colors.sequential.Teal
        )
        # ✅ 수정된 부분: 바 차트에서도 숫자가 명확히 보이도록 크기와 포맷 지정
        fig_bar.update_traces(
            texttemplate='<b>%{text:,.0f} m³</b>', 
            textposition='outside',
            textfont_size=13
        )
        fig_bar.update_layout(
            yaxis={'categoryorder':'total ascending'}, 
            showlegend=False,
            margin=dict(r=100) # 오른쪽에 숫자 표기할 공간(여백) 추가
        )
        st.plotly_chart(fig_bar, use_container_width=True)
