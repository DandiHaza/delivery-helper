import streamlit as st
import pandas as pd
import re
import io
import openpyxl
from datetime import datetime
from zoneinfo import ZoneInfo

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
    '11st_manual': {'key': '11번가', 'skip': 0, 'order': 5},
    'wadiz': {'key': '발송 처리용 주문', 'skip': 0, 'order': 6}
}

def clean_phone(phone):
    if pd.isna(phone): return ""
    return re.sub(r'[^0-9]', '', str(phone))

def identify_product(name):
    name_str = str(name)
    name_upper = name_str.upper()
    name_lower = name_str.lower()
    
    # OH, PH, SH 우선 확인
    if 'OH' in name_upper: return 'OH'
    if 'PH' in name_upper: return 'PH'
    if 'SH' in name_upper: return 'SH'
    
    # 기타 제품 매핑
    if '케이블' in name_str:
        if '스위치' in name_str:
            return '케이블s'
        else:
            return '케이블(일반)'
    if '거치대' in name_str or '휴대폰' in name_str:
        return '휴대폰거치대'
    if '번호판' in name_str or '차량번호' in name_str:
        return '차량번호판'
    if '망치' in name_str or '차량용망치' in name_str:
        return '차량용망치'
    if '도막' in name_str or '측정기' in name_str:
        return '도막측정기'
    
    return name

def get_message(row, cols):
    for col in cols:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            return str(row[col]).strip()
    return ""

def pick_first_col(columns, candidates):
    for col in candidates:
        if col in columns:
            return col
    return None

def format_date(value):
    if pd.isna(value):
        return ""
    try:
        return pd.to_datetime(value).strftime('%Y.%m.%d')
    except Exception:
        return str(value)

def detect_market_by_columns(df):
    cols = set(df.columns.astype(str))

    # 와디즈 감지 (고유 컬럼)
    required_wadiz = {'주문 번호', '주문 상품', '주문 수량', '받는 분'}
    if required_wadiz.issubset(cols):
        return 'wadiz'

    required_11st = {'주문번호', '주소', '상품명', '수량'}
    name_cols_11st = {'수취인', '받는분'}
    phone_cols_11st = {'휴대폰번호', '수취인연락처'}
    if required_11st.issubset(cols) and cols.intersection(name_cols_11st) and cols.intersection(phone_cols_11st):
        return '11st_manual'

    return None

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
        return None

def add_invoice_to_coupang(file_content, file_name, invoice_map):
    """쿠팡 파일에 운송장번호 추가 (서식 유지)"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_content))
        ws = wb.active
        header = [cell.value for cell in ws[1]]
        
        # 주문번호와 운송장번호 컬럼 찾기
        try:
            order_col_idx = header.index('주문번호') + 1
        except:
            return None
        
        # 운송장번호 컬럼이 있는지 확인
        if '운송장번호' in header:
            invoice_col_idx = header.index('운송장번호') + 1
        else:
            # 없으면 맨 끝에 추가
            invoice_col_idx = len(header) + 1
            ws.cell(row=1, column=invoice_col_idx, value='운송장번호')
        
        # 데이터 행에 운송장번호 추가
        for row_idx in range(2, ws.max_row + 1):
            order_no = str(ws.cell(row=row_idx, column=order_col_idx).value)
            invoice = invoice_map.get(order_no, '')
            
            cell = ws.cell(row=row_idx, column=invoice_col_idx)
            cell.value = invoice
            # 숫자를 텍스트로 저장하여 E 표기 방지
            if invoice:
                cell.number_format = '@'  # 텍스트 형식
        
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
        # 파일명으로 매칭되지 않는 경우 컬럼 기반 탐지 시도 (11번가 주문시트 등)
        try:
            df_probe = pd.read_csv(io.BytesIO(content)) if file_name.endswith('.csv') \
                else pd.read_excel(io.BytesIO(content))
            detected = detect_market_by_columns(df_probe)
            if detected:
                market_key = detected
                config = MARKET_CONFIG[detected]
            else:
                # 11번가 주문시트가 상단에 안내 행이 있는 경우를 위한 추가 시도
                df_probe = pd.read_csv(io.BytesIO(content), skiprows=2) if file_name.endswith('.csv') \
                    else pd.read_excel(io.BytesIO(content), skiprows=2)
                detected = detect_market_by_columns(df_probe)
                if detected:
                    market_key = detected
                    config = dict(MARKET_CONFIG[detected])
                    config['skip'] = 2
        except Exception:
            pass

    if market_key == 'unknown':
        return pd.DataFrame()

    try:
        df = pd.read_csv(io.BytesIO(content), skiprows=config.get('skip', 0)) if file_name.endswith('.csv') \
             else pd.read_excel(io.BytesIO(content), skiprows=config.get('skip', 0))

        # 11번가 주문시트는 파일명 매칭이 되더라도 헤더 위치가 다를 수 있어 재시도
        if market_key in ['11st', '11st_manual']:
            required_11st = {'주문번호', '주소', '상품명', '수량'}
            if not required_11st.issubset(set(df.columns.astype(str))):
                df_retry = pd.read_csv(io.BytesIO(content), skiprows=2) if file_name.endswith('.csv') \
                    else pd.read_excel(io.BytesIO(content), skiprows=2)
                if required_11st.issubset(set(df_retry.columns.astype(str))):
                    df = df_retry

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
            phone_col = '휴대폰번호' if '휴대폰번호' in df.columns else (
                '수취인연락처' if '수취인연락처' in df.columns else '전화번호'
            )
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
        elif market_key == 'wadiz':
            df['final_msg'] = df.apply(lambda r: get_message(r, ['배송 요청 사항', '주문 요청 사항']), axis=1)
            mapped = pd.DataFrame({
                '고객주문번호': df['주문 번호'].astype(str),
                '받는분성명': df['받는 분'],
                '받는분전화번호': df['받는 분 연락처'].apply(clean_phone),
                '받는분주소': df['배송지 주소'],
                '배송메세지': df['final_msg'],
                '품목': df['주문 상품'].apply(identify_product),
                '수량': df['주문 수량'],
                '내부정렬키': df['주문 상품'].astype(str)
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
    def sort_key(item):
        order = {'OH': 0, 'PH': 1, 'SH': 2}
        return (order.get(str(item).upper(), 3), str(item))

    formatted = [f"{row['품목']} {int(row['수량'])}개" if row['수량'] > 1 else str(row['품목']) 
                 for _, row in prod_counts.iterrows()]
    formatted.sort(key=lambda x: sort_key(x.split(' ')[0]))

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
if 'order_mgmt_file' not in st.session_state:
    st.session_state.order_mgmt_file = None
if 'order_mgmt_info' not in st.session_state:
    st.session_state.order_mgmt_info = None
if 'order_mgmt_preview' not in st.session_state:
    st.session_state.order_mgmt_preview = None
if 'order_mgmt_raw_data' not in st.session_state:
    st.session_state.order_mgmt_raw_data = None
if 'coupang_delivery_file' not in st.session_state:
    st.session_state.coupang_delivery_file = None
if 'uploaded_market_files' not in st.session_state:
    st.session_state.uploaded_market_files = None

# 사용법 안내
with st.expander("📖 사용법", expanded=False):
    st.markdown("""
    ### 📦 발주 파일 생성
    **파일 준비**
    - **네이버 파일**: 암호가 있는 경우 먼저 제거해주세요
      - 엑셀 파일 열기 → F12 → 도구 → 일반 옵션 → 비밀번호 삭제 → 저장
    
    **사용 순서**
    1. 아래 "📂 파일 업로드"에서 각 마켓의 발주 파일을 업로드하세요 (여러 개 동시 선택 가능)
    2. **파일 생성** 버튼을 클릭하세요
    3. 생성된 파일을 다운로드하세요 (여러 번 가능)
       - `MMDD_HH.xlsx`: CJ택배 업로드용 통합 발주 파일
       - `MMDD_HH_쿠팡_원본정렬.xlsx`: 쿠팡 파일 정렬본 (쿠팡 파일이 있는 경우)
    4. 새로운 파일을 처리하려면 **초기화** 버튼을 누르고 다시 시작하세요
    
    ---
    
    ### 📋 주문관리 시트 생성
    **파일 준비**
    - **CJ택배 파일**: CJ 배송 실적 출력 파일 (운송장번호와 고객주문번호 포함)
    - **마켓 주문 파일**: 각 마켓의 주문 내역 파일 (위에서 업로드한 파일 재사용 가능)
    
    **사용 순서**
    1. "📋 주문관리 시트 생성" 섹션으로 이동하세요
    2. CJ택배 파일을 업로드하세요
    3. 마켓 주문 파일을 업로드하세요
       - 위에서 이미 업로드했다면 "위에서 업로드한 파일 재사용" 체크박스 선택
    4. **주문관리시트 생성** 버튼을 클릭하세요
    5. 생성된 파일을 다운로드하세요
       - `주문관리_MMDD_HH.xlsx`: 송장번호 매칭된 통합 주문 관리 시트
       - `쿠팡발송_MMDD_HH.xlsx`: 쿠팡 발송용 파일 (원본 서식 유지 + 운송장번호)
    6. 데이터 미리보기와 품목별 판매 집계를 확인하세요
    
    **자동 기능**
    - ✅ 같은 주문번호의 제품 자동 통합 (예: OH 2개, PH 1개 → 한 줄로 표시)
    - ✅ CJ 고객주문번호와 매칭되는 송장번호 자동 입력
    - ✅ 발주파일과 동일한 순서로 자동 정렬 (마켓별 → 제품별)
    - ✅ 옥션/지마켓 자동 구분 (주문번호 패턴 분석)
    - ✅ 제품명 자동 분류 (OH, PH, SH, 케이블, 거치대, 번호판 등 9종)
    - ✅ 쿠팡 발송 파일 자동 생성 (원본 서식 유지, 운송장번호 추가)
    
    ---
    
    ### 📌 지원 마켓
    - 네이버 스마트스토어
    - 쿠팡 (DeliveryList)
    - 와디즈 (발송 처리용 주문)
    - 자사몰 (orders)
    - ESM (지마켓/옥션 - 신규주문)
    - 11번가 (allList)
    
    ### 💡 참고사항
    - 파일명 시간 형식: MMDD_HH (예: 0206_15 = 2월 6일 오후 3시)
    - 동일한 배송지로 여러 상품 주문 시 자동 통합
    - 정렬 순서: 네이버→쿠팡→자사몰→ESM→11번가→와디즈 / OH→PH→SH→기타
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
    
    # 세션에 파일 저장 (주문관리시트에서 재사용 가능)
    st.session_state.uploaded_market_files = [(f.name, f.read()) for f in uploaded_files]
    
    # 업로드된 파일 목록 표시
    with st.expander("업로드된 파일 목록"):
        for file in uploaded_files:
            st.write(f"- {file.name}")

if st.button("🚀 발주 파일 생성", type="primary", disabled=not uploaded_files or st.session_state.generated_file is not None):
    with st.spinner("파일 처리 중..."):
        combined_list = []
        coupang_sorted = None
        
        now = datetime.now(ZoneInfo("Asia/Seoul"))
        date_prefix = now.strftime('%m%d')
        time_suffix = now.strftime('%H')
        
        # 파일 처리 (세션에 저장된 파일 사용)
        for file_name, content in st.session_state.uploaded_market_files:
            
            # 쿠팡 파일인 경우 정렬된 버전 생성
            if 'DeliveryList' in file_name:
                coupang_sorted = sort_xlsx_preserving_format(content, '업체상품코드')
            
            # 데이터 처리
            temp_df = process_data(file_name, content)
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
    
    col1, col2 = st.columns(2)
    
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
    
    use_existing = st.checkbox(
        "위에서 업로드한 파일 사용하기",
        value=False,
        disabled=not st.session_state.uploaded_market_files,
        help="발주 파일 생성에서 업로드한 마켓 주문시트를 재사용합니다"
    )
    
    if use_existing and st.session_state.uploaded_market_files:
        st.info(f"✅ {len(st.session_state.uploaded_market_files)}개의 업로드된 파일 사용")
        with st.expander("사용할 파일 목록"):
            for file_name, _ in st.session_state.uploaded_market_files:
                st.write(f"- {file_name}")
        market_files = None

if st.button("🔗 주문관리시트 생성", type="primary", key="gen_order_mgmt"):
    if not cj_file:
        st.error("CJ택배 파일을 업로드해주세요")
    elif not use_existing and not market_files:
        st.error("마켓 주문시트를 업로드하거나 위의 파일을 사용하도록 체크해주세요")
    else:
        with st.spinner("주문관리시트 생성 중..."):
            try:
                # CJ택배 파일 읽기
                cj_content = cj_file.read()
                cj_df = pd.read_csv(io.BytesIO(cj_content)) if cj_file.name.endswith('.csv') \
                    else pd.read_excel(io.BytesIO(cj_content))
                cj_df.columns = cj_df.columns.astype(str).str.strip()
                
                # 운송장번호와 고객주문번호 매핑
                invoice_map = {}
                if '운송장번호' in cj_df.columns and '고객주문번호' in cj_df.columns:
                    for _, row in cj_df.iterrows():
                        order_no = str(row['고객주문번호']).strip()
                        invoice = str(row['운송장번호']).strip()
                        if order_no and invoice and invoice != 'nan':
                            invoice_map[order_no] = invoice
                
                today_str = datetime.now(ZoneInfo("Asia/Seoul")).strftime('%Y.%m.%d')

                # 마켓 주문시트 처리
                all_orders = []
                
                # 사용할 파일 결정
                files_to_process = []
                if use_existing and st.session_state.uploaded_market_files:
                    files_to_process = st.session_state.uploaded_market_files
                else:
                    files_to_process = [(f.name, f.read()) for f in market_files]
                
                for file_name, content in files_to_process:
                    
                    # 마켓별 상세 데이터 추출
                    market_key = 'unknown'
                    config = {}
                    for k, v in MARKET_CONFIG.items():
                        if v['key'] in file_name:
                            market_key = k
                            config = v
                            break
                    
                    # 컬럼 기반 탐지
                    if market_key == 'unknown':
                        try:
                            df_probe = pd.read_csv(io.BytesIO(content)) if file_name.endswith('.csv') \
                                else pd.read_excel(io.BytesIO(content))
                            detected = detect_market_by_columns(df_probe)
                            if detected:
                                market_key = detected
                                config = MARKET_CONFIG[detected]
                            else:
                                df_probe = pd.read_csv(io.BytesIO(content), skiprows=2) if file_name.endswith('.csv') \
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
                    df = pd.read_csv(io.BytesIO(content), skiprows=config.get('skip', 0)) if file_name.endswith('.csv') \
                        else pd.read_excel(io.BytesIO(content), skiprows=config.get('skip', 0))
                    df.columns = df.columns.astype(str).str.strip()
                    
                    # 11번가 헤더 재시도
                    if market_key in ['11st', '11st_manual']:
                        required_11st = {'주문번호', '주소', '상품명', '수량'}
                        if not required_11st.issubset(set(df.columns.astype(str))):
                            df_retry = pd.read_csv(io.BytesIO(content), skiprows=2) if file_name.endswith('.csv') \
                                else pd.read_excel(io.BytesIO(content), skiprows=2)
                            if required_11st.issubset(set(df_retry.columns.astype(str))):
                                df = df_retry
                                df.columns = df.columns.astype(str).str.strip()
                    
                    # 마켓별 데이터 추출
                    channel_name = {'naver': '네이버', 'coupang': '쿠팡', 'own': '자사몰', 'esm': '지마켓', '11st': '11번가', '11st_manual': '11번가'}.get(market_key, '기타')
                    
                    if market_key == 'naver':
                        date_col = pick_first_col(df.columns, ['결제일', '주문일', '결제일시', '주문일시'])
                        buyer_col = pick_first_col(df.columns, ['구매자명', '주문자명', '구매자', '주문자'])
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호']).strip()
                            all_orders.append({
                                '날짜': today_str,
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': identify_product(row.get('상품명', '')),
                                '수량': row.get('수량', ''),
                                '주문인': row.get(buyer_col, '') if buyer_col else '',
                                '수취인': row.get('수취인명', ''),
                                '전화번호': clean_phone(row.get('수취인연락처1', '')),
                                '주소': row.get('통합배송지', ''),
                                '비고': row.get('final_msg', ''),
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key == 'coupang':
                        date_col = pick_first_col(df.columns, ['주문일', '결제완료시각', '결제일시', '주문일시'])
                        buyer_col = pick_first_col(df.columns, ['주문자명', '구매자', '주문자', '구매자명'])
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호']).strip()
                            all_orders.append({
                                '날짜': today_str,
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': identify_product(row.get('등록상품명', '')),
                                '수량': row.get('구매수(수량)', ''),
                                '주문인': row.get(buyer_col, '') if buyer_col else '',
                                '수취인': row.get('수취인이름', ''),
                                '전화번호': clean_phone(row.get('수취인전화번호', '')),
                                '주소': row.get('수취인 주소', ''),
                                '비고': row.get('final_msg', ''),
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key == 'esm':
                        date_col = pick_first_col(df.columns, ['결제일시', '주문일', '결제일', '주문일시'])
                        buyer_col = pick_first_col(df.columns, ['주문자명', '구매자명', '주문자', '구매자'])
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송시 요구사항', '배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호']).strip()
                            
                            # 주문번호 패턴으로 옥션/지마켓 구분
                            if len(order_no) == 10:
                                if order_no.startswith('2'):
                                    actual_channel = '옥션'
                                elif order_no.startswith('4'):
                                    actual_channel = '지마켓'
                                else:
                                    actual_channel = channel_name
                            else:
                                actual_channel = channel_name
                            
                            all_orders.append({
                                '날짜': today_str,
                                '채널': actual_channel,
                                '주문번호': order_no,
                                '상품명': identify_product(row.get('상품명', '')),
                                '수량': row.get('수량', ''),
                                '주문인': row.get(buyer_col, '') if buyer_col else '',
                                '수취인': row.get('수령인명', ''),
                                '전화번호': clean_phone(row.get('수령인 휴대폰', '')),
                                '주소': row.get('주소', ''),
                                '비고': row.get('final_msg', ''),
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key in ['11st', '11st_manual']:
                        date_col = pick_first_col(df.columns, ['결제일시', '주문일', '결제일', '주문일시'])
                        buyer_col = pick_first_col(df.columns, ['구매자', '주문자', '구매자명', '주문자명'])
                        name_col = pick_first_col(df.columns, ['수취인', '받는분'])
                        phone_col = pick_first_col(df.columns, ['휴대폰번호', '수취인연락처', '전화번호'])
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['배송메시지', '배송메세지', '비고']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호']).strip()
                            all_orders.append({
                                '날짜': today_str,
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': identify_product(row.get('상품명', '')),
                                '수량': row.get('수량', ''),
                                '주문인': row.get(buyer_col, '') if buyer_col else '',
                                '수취인': row.get(name_col, '') if name_col else '',
                                '전화번호': clean_phone(row.get(phone_col, '')) if phone_col else '',
                                '주소': row.get('주소', ''),
                                '비고': row.get('final_msg', ''),
                                '송장번호': invoice_map.get(order_no, '')
                            })
                    
                    elif market_key == 'own':
                        date_col = pick_first_col(df.columns, ['주문일시', '주문일', '결제일', '결제일시'])
                        buyer_col = pick_first_col(df.columns, ['주문자', '구매자', '주문자명', '구매자명'])
                        df['final_msg'] = df.apply(lambda r: get_message(r, ['비고', '배송메세지']), axis=1)
                        
                        for _, row in df.iterrows():
                            order_no = str(row['주문번호']).strip()
                            all_orders.append({
                                '날짜': today_str,
                                '채널': channel_name,
                                '주문번호': order_no,
                                '상품명': identify_product(row.get('주문상품명', '')),
                                '수량': row.get('수량', ''),
                                '주문인': row.get(buyer_col, '') if buyer_col else '',
                                '수취인': row.get('수령인', ''),
                                '전화번호': clean_phone(row.get('핸드폰', '')),
                                '주소': row.get('주소', ''),
                                '비고': row.get('final_msg', ''),
                                '송장번호': invoice_map.get(order_no, '')
                            })
                
                # 주문관리시트 생성
                if all_orders:
                    mgmt_df = pd.DataFrame(all_orders)
                    
                    # 같은 주문번호로 제품 통합
                    consolidated_list = []
                    
                    for (channel, order_no), group in mgmt_df.groupby(['채널', '주문번호']):
                        # 제품별 수량 집계
                        prod_counts = {}
                        for _, row in group.iterrows():
                            prod = row['상품명']
                            qty = row['수량']
                            if prod in prod_counts:
                                prod_counts[prod] += qty
                            else:
                                prod_counts[prod] = qty
                        
                        # OH, PH, SH 순서로 정렬
                        def get_sort_priority(prod_name):
                            prod_upper = str(prod_name).strip().upper()
                            if prod_upper == 'OH':
                                return (0, prod_name)
                            elif prod_upper == 'PH':
                                return (1, prod_name)
                            elif prod_upper == 'SH':
                                return (2, prod_name)
                            else:
                                return (3, prod_name)
                        
                        sorted_prods = sorted(prod_counts.items(), key=lambda x: get_sort_priority(x[0]))
                        
                        # "OH 2개, PH 1개" 형태로 포맷팅
                        formatted = []
                        for prod, qty in sorted_prods:
                            if qty > 1:
                                formatted.append(f"{prod} {int(qty)}개")
                            else:
                                formatted.append(str(prod))
                        
                        # 첫 번째 제품으로 정렬키 결정
                        first_prod = sorted_prods[0][0] if sorted_prods else ''
                        first_prod_priority = get_sort_priority(first_prod)[0]
                        
                        # 마켓 순서 매핑
                        market_order_map = {'네이버': 1, '쿠팡': 2, '자사몰': 3, '지마켓': 4, '11번가': 5}
                        market_order = market_order_map.get(channel, 99)
                        
                        consolidated_list.append({
                            '날짜': group.iloc[0]['날짜'],
                            '채널': channel,
                            '주문번호': order_no,
                            '상품명': ", ".join(formatted),
                            '수량': int(group['수량'].sum()),
                            '주문인': group.iloc[0]['주문인'],
                            '수취인': group.iloc[0]['수취인'],
                            '전화번호': group.iloc[0]['전화번호'],
                            '주소': group.iloc[0]['주소'],
                            '비고': group.iloc[0]['비고'],
                            '송장번호': group.iloc[0]['송장번호'],
                            '마켓순서': market_order,
                            '상품순서': first_prod_priority
                        })
                    
                    consolidated = pd.DataFrame(consolidated_list)
                    # 발주파일과 같은 순서로 정렬: 마켓 → 상품
                    consolidated = consolidated.sort_values(by=['마켓순서', '상품순서'])
                    # 정렬용 컬럼 제거
                    consolidated = consolidated.drop(columns=['마켓순서', '상품순서'])
                    
                    # 쿠팡 발송 파일 생성
                    coupang_delivery = None
                    if use_existing and st.session_state.uploaded_market_files:
                        # 업로드된 파일에서 쿠팡 파일 찾기
                        for file_name, content in st.session_state.uploaded_market_files:
                            if 'DeliveryList' in file_name:
                                coupang_delivery = add_invoice_to_coupang(content, file_name, invoice_map)
                                break
                    elif market_files:
                        # 새로 업로드한 파일에서 쿠팡 파일 찾기
                        for f in market_files:
                            if 'DeliveryList' in f.name:
                                content = f.read()
                                coupang_delivery = add_invoice_to_coupang(content, f.name, invoice_map)
                                break
                    
                    # 엑셀 파일 생성
                    output = io.BytesIO()
                    consolidated.to_excel(output, index=False)
                    output.seek(0)
                    
                    now = datetime.now(ZoneInfo("Asia/Seoul"))
                    filename = f"주문관리_{now.strftime('%m%d_%H')}.xlsx"
                    
                    st.session_state.order_mgmt_file = output.getvalue()
                    st.session_state.order_mgmt_info = {
                        'filename': filename,
                        'count': len(consolidated),
                        'matched': len(consolidated[consolidated['송장번호'] != ''])
                    }
                    st.session_state.order_mgmt_preview = consolidated
                    st.session_state.order_mgmt_raw_data = all_orders
                    st.session_state.coupang_delivery_file = coupang_delivery
                    
                    st.success("✅ 주문관리시트 생성 완료!")
                    st.rerun()
                else:
                    st.error("❌ 처리할 수 있는 주문 데이터가 없습니다.")
                    
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")

# 주문관리시트 다운로드
if st.session_state.order_mgmt_file:
    st.markdown("### 📥 주문관리시트 다운로드")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📋 주문관리시트 다운로드",
            data=st.session_state.order_mgmt_file,
            file_name=st.session_state.order_mgmt_info['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    if st.session_state.coupang_delivery_file:
        with col2:
            now = datetime.now(ZoneInfo("Asia/Seoul"))
            coupang_filename = f"쿠팡발송_{now.strftime('%m%d_%H')}.xlsx"
            st.download_button(
                label="📦 쿠팡 발송 파일 다운로드",
                data=st.session_state.coupang_delivery_file,
                file_name=coupang_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    st.info(f"총 {st.session_state.order_mgmt_info['count']}건 | 송장번호 매칭 {st.session_state.order_mgmt_info['matched']}건")
    
    # 미리보기
    with st.expander("📊 데이터 미리보기", expanded=True):
        st.dataframe(st.session_state.order_mgmt_preview, use_container_width=True)
    
    # 품목별 판매 집계
    if st.session_state.order_mgmt_raw_data:
        with st.expander("📈 품목별 판매 집계", expanded=False):
            raw_df = pd.DataFrame(st.session_state.order_mgmt_raw_data)
            product_summary = raw_df.groupby('상품명')['수량'].sum().reset_index()
            product_summary.columns = ['품목', '판매 수량']
            
            # 품목 순서 정의
            product_order = {
                'OH': 0,
                'PH': 1,
                'SH': 2,
                '케이블(일반)': 3,
                '케이블s': 4,
                '휴대폰거치대': 5,
                '차량번호판': 6,
                '차량용망치': 7,
                '도막측정기': 8
            }
            
            # 정렬키 추가
            product_summary['순서'] = product_summary['품목'].map(lambda x: product_order.get(x, 99))
            product_summary = product_summary.sort_values(by='순서')
            product_summary = product_summary[['품목', '판매 수량']]
            
            st.dataframe(product_summary, use_container_width=True, hide_index=True)
            st.info(f"총 품목 수: {len(product_summary)}개")
    
    if st.button("🔄 새 주문관리시트 생성", key="reset_mgmt"):
        st.session_state.order_mgmt_file = None
        st.session_state.order_mgmt_info = None
        st.session_state.order_mgmt_preview = None
        st.session_state.order_mgmt_raw_data = None
        st.session_state.coupang_delivery_file = None
        st.rerun()

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
    자동 발주 파일 생성기 | Made by 🦖 DandiHaza
    </div>
    """,
    unsafe_allow_html=True
)
