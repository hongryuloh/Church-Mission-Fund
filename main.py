# 4. 인쇄용 집계표 (Admin 전용)
elif menu == "🖨️ 인쇄용 집계표":
    st.subheader("🖨️ 인쇄용 엑셀 다운로드")
    months = sorted(list(set([f"{TARGET_YEAR}{str(m).zfill(2)}" for m in range(1, 13)] + list(df_income['년월'].unique()))), reverse=True)
    
    # 💡 [기능 추가] 현재 날짜를 기준으로 '전월'을 계산하여 기본값으로 설정
    now = datetime.now()
    prev_month = now.month - 1
    calc_year = TARGET_YEAR
    
    # 1월인 경우 작년 12월을 가리키도록 처리
    if prev_month == 0:
        prev_month = 12
        calc_year -= 1
        
    default_month_str = f"{calc_year}{prev_month:02d}"
    
    # 계산된 '전월'이 콤보박스 목록의 몇 번째에 있는지 찾기 (없으면 0번째)
    default_idx = months.index(default_month_str) if default_month_str in months else 0
    
    # 콤보박스에 index 속성을 추가하여 기본 선택값 지정
    target_month = st.selectbox("📌 기준월 선택", months, index=default_idx)
    
    if st.button("🔄 인쇄용 파일 생성", use_container_width=True):
        with st.spinner("엑셀 파일을 생성 중입니다..."):
            donors = set(df_income[df_income['년월'] == target_month]['이름'].apply(lambda x: str(x).strip()).unique())
            with conn.session as s:
                for idx, row in df_target.iterrows():
                    name = str(row.get('이름')).strip()
                    if not name or name == 'nan' or name == '합계': continue
                    print_val = 'Y' if name in donors else 'N'
                    s.execute(text("UPDATE target SET print_yn=:p WHERE name=:n AND target_year=:y"), {"p": print_val, "n": name, "y": TARGET_YEAR})
                s.commit()
            clear_db_cache(); _, df_target_updated, _, _ = load_data(TARGET_YEAR)
            st.session_state.download_data = generate_summary_excel(df_income, df_target_updated, target_month)
            st.success(f"✅ 완성되었습니다!")
            
    if 'download_data' in st.session_state:
        st.download_button("📥 인쇄용 엑셀 다운로드", data=st.session_state.download_data, file_name=f"선교헌금_인쇄용_{target_month}.xlsx", use_container_width=True)
