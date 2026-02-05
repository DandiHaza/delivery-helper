# 주문관리시트 생성 섹션 추가 코드

# Footer 앞에 추가할 내용:

st.markdown("---")
st.markdown("## 📋 주문관리시트 생성 (송장번호 매칭)")
st.markdown("발주 후 CJ택배에서 받은 송장번호 파일과 마켓 주문시트를 매칭하여 주문관리시트를 생성합니다.")

# 세션 상태 초기화 (주문관리시트용)
if 'order_mgmt_file' not in st.session_state:
    st.session_state.order_mgmt_file = None
if 'order_mgmt_info' not in st.session_state:
    st.session_state.order_mgmt_info = None

col_a, col_b = st.columns(2)

with col_a:
    cj_file = st.file_uploader(
        "CJ택배 출력 파일 업로드",
        type=['xlsx', 'xls', 'csv'],
        key="cj_upload",
        help="운송장번호와 고객주문번호가 포함된 CJ택배 출력 파일"
    )

with col_b:
    market_files = st.file_uploader(
        "마켓 주문시트 업로드",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        key="market_upload",
        help="네이버, 쿠팡, 11번가 등 마켓 주문시트"
    )

if st.button("🔗 주문관리시트 생성", type="primary", key="gen_order_mgmt"):
    if not cj_file:
        st.error("CJ택배 파일을 업로드해주세요")
    elif not market_files:
        st.error("마켓 주문시트를 업로드해주세요")
    else:
        with st.spinner("주문관리시트 생성 중..."):
            try:
                # CJ택배 파일 읽기
                cj_content = cj_file.read()
                cj_df = pd.read_csv(io.BytesIO(cj_content)) if cj_file.name.endswith('.csv') \
                    else pd.read_excel(io.BytesIO(cj_content))
                
                # 운송장번호와 고객주문번호 매핑
                invoice_map = {}
                if '운송장번호' in cj_df.columns and '고객주문번호' in cj_df.columns:
                    for _, row in cj_df.iterrows():
                        order_no = str(row['고객주문번호']).strip()
                        invoice = str(row['운송장번호']).strip()
                        if order_no and invoice and invoice != 'nan':
                            invoice_map[order_no] = invoice
                
                # 마켓 주문시트 처리
                all_orders = []
                for market_file in market_files:
                    content = market_file.read()
                    
                    # 마켓별 상세 데이터 추출
                    market_key = 'unknown'
                    config = {}
                    for k, v in MARKET_CONFIG.items():
                        if v['key'] in market_file.name:
                            market_key = k
                            config = v
                            break
                    
                    # 컬럼 기반 탐지
                    if market_key == 'unknown':
                        try:
                            df_probe = pd.read_csv(io.BytesIO(content)) if market_file.name.endswith('.csv') \
                                else pd.read_excel(io.BytesIO(content))
                            detected = detect_market_by_columns(df_probe)
                            if detected:
                                market_key = detected
                                config = MARKET_CONFIG[detected]
                            else:
                                df_probe = pd.read_csv(io.BytesIO(content), skiprows=2) if market_file.name.endswith('.csv') \
                                    else pd.read_excel(io.BytesIO(content), skiprows=2)
                                detected = detect_market_by_columns(df_probe)
                                if detected:
                                    market_key = detected
                                    config = dict(MARKET_CONFIG[detected])
                                    config['skip'] = 2
                        except Exception:
                            pass
                    
                    if market_key == 'unknown':
                        continue
                    
                    # 데이터 읽기
                    df = pd.read_csv(io.BytesIO(content), skiprows=config.get('skip', 0)) if market_file.name.endswith('.csv') \
                         else pd.read_excel(io.BytesIO(content), skiprows=config.get('skip', 0))
                    
                    # 11번가 헤더 재시도
                    if market_key in ['11st', '11st_manual']:
                        required_11st = {'주문번호', '주소', '상품명', '수량'}
                        if not required_11st.issubset(set(df.columns.astype(str))):
                            df_retry = pd.read_csv(io.BytesIO(content), skiprows=2) if market_file.name.endswith('.csv') \
                                else pd.read_excel(io.BytesIO(content), skiprows=2)
                            if required_11st.issubset(set(df_retry.columns.astype(str))):
                                df = df_retry
                    
                    # 마켓별 데이터 추출
                    channel_name = {'naver': '스마트스토어', 'coupang': '쿠팡', 'own': '자사몰', 'esm': '지마켓', '11st': '11번가', '11st_manual': '11번가'}.get(market_key, '기타')
                    
                    if market_key == 'naver':
                        date_col = '결제일' if '결제일' in df.columns else '주문일'
                        buyer_col = '구매자명' if '구매자명' in df.columns else '주문자명'
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호'])
                            all_orders.append({
                                '날짜': pd.to_datetime(row[date_col]).strftime('%Y.%m.%d') if pd.notna(row[date_col]) else '',
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': row['상품명'],
                                '수량': row['수량'],
                                '주문인': row[buyer_col] if buyer_col in df.columns else '',
                                '수취인': row['수취인명'],
                                '전화번호': clean_phone(row['수취인연락처1']),
                                '주소': row['통합배송지'],
                                '비고': row['final_msg'],
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key == 'coupang':
                        date_col = '주문일' if '주문일' in df.columns else '결제완료시각'
                        buyer_col = '주문자명' if '주문자명' in df.columns else '구매자'
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호'])
                            all_orders.append({
                                '날짜': pd.to_datetime(row[date_col]).strftime('%Y.%m.%d') if pd.notna(row[date_col]) else '',
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': row['등록상품명'],
                                '수량': row['구매수(수량)'],
                                '주문인': row[buyer_col] if buyer_col in df.columns else '',
                                '수취인': row['수취인이름'],
                                '전화번호': clean_phone(row['수취인전화번호']),
                                '주소': row['수취인 주소'],
                                '비고': row['final_msg'],
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key == 'esm':
                        date_col = '결제일시' if '결제일시' in df.columns else '주문일'
                        buyer_col = '주문자명' if '주문자명' in df.columns else '구매자명'
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송시 요구사항', '배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호'])
                            all_orders.append({
                                '날짜': pd.to_datetime(row[date_col]).strftime('%Y.%m.%d') if pd.notna(row[date_col]) else '',
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': row['상품명'],
                                '수량': row['수량'],
                                '주문인': row[buyer_col] if buyer_col in df.columns else '',
                                '수취인': row['수령인명'],
                                '전화번호': clean_phone(row['수령인 휴대폰']),
                                '주소': row['주소'],
                                '비고': row['final_msg'],
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key in ['11st', '11st_manual']:
                        date_col = '결제일시' if '결제일시' in df.columns else '주문일'
                        buyer_col = '구매자' if '구매자' in df.columns else '주문자'
                        name_col = '수취인' if '수취인' in df.columns else '받는분'
                        phone_col = '휴대폰번호' if '휴대폰번호' in df.columns else (
                            '수취인연락처' if '수취인연락처' in df.columns else '전화번호'
                        )
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메시지', '배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호'])
                            all_orders.append({
                                '날짜': pd.to_datetime(row[date_col]).strftime('%Y.%m.%d') if pd.notna(row[date_col]) else '',
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': row['상품명'],
                                '수량': row['수량'],
                                '주문인': row[buyer_col] if buyer_col in df.columns else '',
                                '수취인': row[name_col],
                                '전화번호': clean_phone(row[phone_col]),
                                '주소': row['주소'],
                                '비고': row['final_msg'],
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key == 'own':
                        date_col = '주문일시' if '주문일시' in df.columns else '주문일'
                        buyer_col = '주문자' if '주문자' in df.columns else '구매자'
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['비고', '배송메세지']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호'])
                            all_orders.append({
                                '날짜': pd.to_datetime(row[date_col]).strftime('%Y.%m.%d') if pd.notna(row[date_col]) else '',
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': row['주문상품명'],
                                '수량': row['수량'],
                                '주문인': row[buyer_col] if buyer_col in df.columns else '',
                                '수취인': row['수령인'],
                                '전화번호': clean_phone(row['핸드폰']),
                                '주소': row['주소'],
                                '비고': row['final_msg'],
                                '송장번호': invoice_map.get(order_no, '')
                            })
                
                # 주문관리시트 생성
                if all_orders:
                    mgmt_df = pd.DataFrame(all_orders)
                    mgmt_df = mgmt_df.sort_values(by=['날짜', '채널'])
                    
                    # 엑셀 파일 생성
                    output = io.BytesIO()
                    mgmt_df.to_excel(output, index=False)
                    output.seek(0)
                    
                    now = datetime.now(ZoneInfo("Asia/Seoul"))
                    filename = f"주문관리_{now.strftime('%Y%m%d')}.xlsx"
                    
                    st.session_state.order_mgmt_file = output.getvalue()
                    st.session_state.order_mgmt_info = {
                        'filename': filename,
                        'count': len(mgmt_df),
                        'matched': len([o for o in all_orders if o['송장번호']])
                    }
                    
                    st.success("✅ 주문관리시트 생성 완료!")
                    st.rerun()
                else:
                    st.error("❌ 처리할 수 있는 주문 데이터가 없습니다.")
                    
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")

# 주문관리시트 다운로드
if st.session_state.order_mgmt_file:
    st.markdown("### 📥 주문관리시트 다운로드")
    st.download_button(
        label="📋 주문관리시트 다운로드",
        data=st.session_state.order_mgmt_file,
        file_name=st.session_state.order_mgmt_info['filename'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.info(f"총 {st.session_state.order_mgmt_info['count']}건 | 송장번호 매칭 {st.session_state.order_mgmt_info['matched']}건")
    
    if st.button("🔄 새 주문관리시트 생성", key="reset_mgmt"):
        st.session_state.order_mgmt_file = None
        st.session_state.order_mgmt_info = None
        st.rerun()
