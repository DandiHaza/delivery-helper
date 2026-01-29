import streamlit as st
import pandas as pd
import re
import io
import openpyxl
from datetime import datetime

# 페이지 설정
st.set_page_config(
    page_title="자동 발주 파일 생성기",
    page_icon="📦",
    layout="wide"
)

# ==========================================
# 설정 및 함수들
# ==========================================
MARKET_CONFIG = {
    'naver': {'key': '스마트스토어', 'skip': 1, 'order': 1},
    'coupang': {'key': 'DeliveryList', 'skip': 0, 'order': 2},
    'own': {'key': 'orders', 'skip': 0, 'order': 3},
    'esm': {'key': '신규주문', 'skip': 0, 'order': 4},
    '11st': {'key': 'allList', 'skip': 2, 'order': 5},
    '11st_manual': {'key': '11번가', 'skip': 0, 'order': 5}
}

def clean_phone(phone):
    if pd.isna(phone): return ""
    return re.sub(r'[^0-9]', '', str(phone))

def identify_product(name):
    name_upper = str(name).upper()
    if 'OH' in name_upper: return 'OH'
    if 'PH' in name_upper: return 'PH'
    if 'SH' in name_upper: return 'SH'
    return name

def get_message(row, cols):
    for col in cols:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            return str(row[col]).strip()
    return ""

def sort_xlsx_preserving_format(file_content, target_col_name):
    """원본 서식을 유지하며 업체상품코드 기준으로 정렬"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_content))
        ws = wb.active
        header = [cell.value for cell in ws[1]]
        
        try:
            col_idx = header.index(target_col_name)
        except:
            return None

        rows = list(ws.iter_rows(min_row=2, values_only=False))
        rows.sort(key=lambda x: str(x[col_idx].value) if x[col_idx].value is not None else "")

        data_styles = []
        for row in rows:
            data_styles.append([(cell.value, cell._style) for cell in row])

        ws.delete_rows(2, ws.max_row)
        for r_idx, row_data in enumerate(data_styles, start=2):
            for c_idx, (val, style) in enumerate(row_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if style:
                    cell._style = style
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()
    except Exception as e:
        st.warning(f"쿠팡 정렬 중 오류: {e}")
        return None

def process_data(file_name, content):
    market_key = 'unknown'
    config = {}
    for k, v in MARKET_CONFIG.items():
        if v['key'] in file_name:
            market_key = k
            config = v
            break

    if market_key == 'unknown':
        return pd.DataFrame()

    try:
        df = pd.read_csv(io.BytesIO(content), skiprows=config.get('skip', 0)) if file_name.endswith('.csv') \
             else pd.read_excel(io.BytesIO(content), skiprows=config.get('skip', 0))

        if market_key == 'naver':
            df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메세지', '비고']), axis=1)
            mapped = pd.DataFrame({
                '고객주문번호': df['주문번호'].astype(str),
                '받는분성명': df['수취인명'],
                '받는분전화번호': df['수취인연락처1'].apply(clean_phone),
                '받는분주소': df['통합배송지'],
                '배송메세지': df['final_msg'],
                '품목': df['상품명'].apply(identify_product),
                '수량': df['수량'],
                '내부정렬키': df['상품명'].astype(str)
            })
        elif market_key == 'coupang':
            df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메세지', '비고']), axis=1)
            mapped = pd.DataFrame({
                '고객주문번호': df['주문번호'].astype(str),
                '받는분성명': df['수취인이름'],
                '받는분전화번호': df['수취인전화번호'].apply(clean_phone),
                '받는분주소': df['수취인 주소'],
                '배송메세지': df['final_msg'],
                '품목': df['등록상품명'].apply(identify_product),
                '수량': df['구매수(수량)'],
                '내부정렬키': df['업체상품코드'].astype(str)
            })
        elif market_key == 'esm':
            df['final_msg'] = df.apply(lambda r: get_message(r, ['배송시 요구사항', '배송메세지', '비고']), axis=1)
            mapped = pd.DataFrame({
                '고객주문번호': df['주문번호'].astype(str),
                '받는분성명': df['수령인명'],
                '받는분전화번호': df['수령인 휴대폰'].apply(clean_phone),
                '받는분주소': df['주소'],
                '배송메세지': df['final_msg'],
                '품목': df['상품명'].apply(identify_product),
                '수량': df['수량'],
                '내부정렬키': df['상품명'].astype(str)
            })
        elif market_key in ['11st', '11st_manual']:
            name_col = '수취인' if '수취인' in df.columns else '받는분'
            phone_col = '휴대폰번호' if '휴대폰번호' in df.columns else '수취인연락처'
            df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메시지', '배송메세지', '비고']), axis=1)
            mapped = pd.DataFrame({
                '고객주문번호': df['주문번호'].astype(str),
                '받는분성명': df[name_col],
                '받는분전화번호': df[phone_col].apply(clean_phone),
                '받는분주소': df['주소'],
                '배송메세지': df['final_msg'],
                '품목': df['상품명'].apply(identify_product),
                '수량': df['수량'],
                '내부정렬키': df['상품명'].astype(str)
            })
        elif market_key == 'own':
            df['final_msg'] = df.apply(lambda r: get_message(r, ['비고', '배송메세지']), axis=1)
            mapped = pd.DataFrame({
                '고객주문번호': df['주문번호'].astype(str),
                '받는분성명': df['수령인'],
                '받는분전화번호': df['핸드폰'].apply(clean_phone),
                '받는분주소': df['주소'],
                '배송메세지': df['final_msg'],
                '품목': df['주문상품명'].apply(identify_product),
                '수량': df['수량'],
                '내부정렬키': df['주문상품명'].astype(str)
            })
        else:
            return pd.DataFrame()

        mapped['마켓순서'] = config['order']
        return mapped
    except Exception as e:
        st.error(f"❌ {file_name} 처리 실패: {e}")
        return pd.DataFrame()

def consolidate(group):
    prod_counts = group.groupby('품목')['수량'].sum().reset_index()
    formatted = [f"{row['품목']} {int(row['수량'])}개" if row['수량'] > 1 else str(row['품목']) 
                 for _, row in prod_counts.iterrows()]
    formatted.sort()

    non_empty_msgs = group['배송메세지'][group['배송메세지'] != ""].unique()
    final_msg = non_empty_msgs[0] if len(non_empty_msgs) > 0 else ""

    return {
        '고객주문번호': group.iloc[0]['고객주문번호'],
        '받는분성명': group.iloc[0]['받는분성명'],
        '받는분전화번호': group.iloc[0]['받는분전화번호'],
        '받는분주소': group.iloc[0]['받는분주소'],
        '배송메세지': final_msg,
        '품목명': ", ".join(formatted),
        '기타1': group['수량'].sum(),
        '마켓순서': group.iloc[0]['마켓순서'],
        '최종정렬키': group['내부정렬키'].min()
    }

# ==========================================
# Streamlit UI
# ==========================================
st.title("📦 자동 발주 파일 생성기")
st.markdown("---")

# 세션 상태 초기화
if 'generated_file' not in st.session_state:
    st.session_state.generated_file = None
if 'coupang_file' not in st.session_state:
    st.session_state.coupang_file = None
if 'file_info' not in st.session_state:
    st.session_state.file_info = None
if 'preview_data' not in st.session_state:
    st.session_state.preview_data = None

# 사용법 안내
with st.expander("📖 사용법", expanded=False):
    st.markdown("""
    ### 사용 방법
    1. **네이버 파일**: 암호가 있는 경우 먼저 제거해주세요
       - 엑셀 파일 열기 → F12 → 도구 → 일반 옵션 → 비밀번호 삭제 → 저장
    2. 아래에서 각 마켓의 발주 파일을 업로드하세요
    3. **파일 생성** 버튼을 클릭하세요
    4. 생성된 파일을 원하는 만큼 다운로드하세요 (여러 번 가능)
    5. 새로운 파일을 처리하려면 **초기화** 버튼을 누르고 다시 시작하세요
    
    ### 지원 마켓
    - 네이버 스마트스토어
    - 쿠팡 (DeliveryList)
    - 자사몰 (orders)
    - ESM (지마켓/옥션 - 신규주문)
    - 11번가 (allList)
    """)

st.markdown("### 📂 파일 업로드")

# 초기화 버튼 (생성된 파일이 있을 때만 표시)
if st.session_state.generated_file:
    if st.button("🔄 초기화 (새로운 파일 처리)", type="secondary"):
        st.session_state.generated_file = None
        st.session_state.coupang_file = None
        st.session_state.file_info = None
        st.session_state.preview_data = None
        st.rerun()

uploaded_files = st.file_uploader(
    "발주 파일을 선택하세요 (여러 파일 선택 가능)",
    type=['csv', 'xlsx', 'xls'],
    accept_multiple_files=True,
    help="네이버, 쿠팡, 자사몰, ESM, 11번가 등의 발주 파일을 모두 선택하세요",
    disabled=st.session_state.generated_file is not None
)

if uploaded_files and not st.session_state.generated_file:
    st.success(f"✅ {len(uploaded_files)}개 파일 업로드됨")
    
    # 업로드된 파일 목록 표시
    with st.expander("업로드된 파일 목록"):
        for file in uploaded_files:
            st.write(f"- {file.name}")

if st.button("🚀 발주 파일 생성", type="primary", disabled=not uploaded_files or st.session_state.generated_file is not None):
    with st.spinner("파일 처리 중..."):
        combined_list = []
        coupang_sorted = None
        
        now = datetime.now()
        date_prefix = now.strftime('%m%d')
        time_suffix = '09' if now.hour < 12 else '16'
        
        # 파일 처리
        for uploaded_file in uploaded_files:
            content = uploaded_file.read()
            
            # 쿠팡 파일인 경우 정렬된 버전 생성
            if 'DeliveryList' in uploaded_file.name:
                coupang_sorted = sort_xlsx_preserving_format(content, '업체상품코드')
            
            # 데이터 처리
            temp_df = process_data(uploaded_file.name, content)
            if not temp_df.empty:
                combined_list.append(temp_df)
        
        if combined_list:
            # 데이터 병합 및 처리
            full_df = pd.concat(combined_list, ignore_index=True)
            
            final_data = []
            groups = full_df.groupby(['받는분성명', '받는분전화번호', '받는분주소'], sort=False)
            for name, group in groups:
                final_data.append(consolidate(group))
            
            final_df = pd.DataFrame(final_data)
            final_df = final_df.sort_values(by=['마켓순서', '최종정렬키'])
            
            # 최종 파일 생성
            final_filename = f"{date_prefix}_{time_suffix}.xlsx"
            final_cols = ['고객주문번호', '받는분성명', '받는분전화번호', '받는분주소(전체, 분할)', '배송메세지1', '품목명', '기타1']
            
            output = io.BytesIO()
            final_df.rename(columns={
                '받는분주소': '받는분주소(전체, 분할)',
                '배송메세지': '배송메세지1'
            }).to_excel(output, index=False, columns=final_cols)
            output.seek(0)
            
            # 세션 상태에 저장
            st.session_state.generated_file = output.getvalue()
            st.session_state.coupang_file = coupang_sorted
            st.session_state.file_info = {
                'filename': final_filename,
                'coupang_filename': f"{date_prefix}_{time_suffix}_쿠팡_원본정렬.xlsx",
                'order_count': len(final_df)
            }
            st.session_state.preview_data = final_df[['고객주문번호', '받는분성명', '품목명', '기타1']]
            
            st.success("✅ 발주 파일 생성 완료!")
            st.rerun()
        else:
            st.error("❌ 처리할 수 있는 파일이 없습니다. 파일 형식을 확인해주세요.")

# 생성된 파일이 있으면 다운로드 섹션 표시
if st.session_state.generated_file:
    st.markdown("---")
    st.markdown("### 📥 파일 다운로드")
    st.info("💡 아래 버튼을 원하는 만큼 클릭하여 파일을 다운로드하세요. 다운로드 후에도 파일은 유지됩니다.")
    
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        st.download_button(
            label="📄 발주 파일 다운로드",
            data=st.session_state.generated_file,
            file_name=st.session_state.file_info['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    if st.session_state.coupang_file:
        with col2:
            st.download_button(
                label="📄 쿠팡 정렬 파일 다운로드",
                data=st.session_state.coupang_file,
                file_name=st.session_state.file_info['coupang_filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # 미리보기
    with st.expander("📊 데이터 미리보기", expanded=True):
        st.dataframe(st.session_state.preview_data, use_container_width=True)
        st.info(f"총 주문 건수: {st.session_state.file_info['order_count']}건")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
    자동 발주 파일 생성기 | Made with ❤️ using Streamlit
    </div>
    """,
    unsafe_allow_html=True
)
