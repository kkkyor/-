import streamlit as st
import pandas as pd
import gspread
from datetime import datetime
import re
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
import urllib.parse
import fitz
from PIL import Image
import io

# --------------------------------------------------------------------------
# 1. Google Sheets 연동 및 데이터 처리 함수
# --------------------------------------------------------------------------

# PDF를 이미지로 변환하는 함수 (새로 추가)
# PDF를 이미지로 변환하는 함수 (기본 페이지 변경)
def convert_pdf_page_to_image(pdf_bytes, page_number=1): # ◀️ 이 숫자를 0에서 1로 변경
    """PDF 파일의 특정 페이지를 이미지 객체로 변환합니다."""
    try:
        # 바이트 데이터로부터 PDF 문서 열기
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # 페이지 수가 요청된 페이지 번호보다 적은 경우 처리
        if len(pdf_document) <= page_number:
            st.warning(f"'{page_number + 1}'번째 페이지가 존재하지 않아 첫 페이지를 표시합니다.")
            page_number = 0
            if len(pdf_document) == 0:
                st.error("PDF에 페이지가 없습니다.")
                return None

        # 지정된 페이지 선택 (0은 첫 페이지, 1은 두 번째 페이지)
        page = pdf_document.load_page(page_number)
        
        # 페이지를 이미지(pixmap)로 렌더링
        pix = page.get_pixmap()
        
        # pixmap을 이미지 바이트로 변환
        img_bytes = pix.tobytes("png")
        
        # 바이트 데이터로부터 Pillow 이미지 객체 생성
        image = Image.open(io.BytesIO(img_bytes))
        return image
    except Exception as e:
        # 오류 발생 시 None 반환
        st.error(f"PDF를 이미지로 변환하는 중 오류 발생: {e}")
        return None

def connect_to_sheet():
    """Google Sheets에 연결하고 워크시트 객체를 반환합니다."""
    try:
        gc = gspread.service_account(filename='credentials.json')
        # gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        spreadsheet = gc.open("계약관리DB") # 실제 스프레드시트 이름으로 변경
        worksheet = spreadsheet.sheet1
        return worksheet
    except Exception as e:
        st.error(f"Google Sheets 연결에 실패했습니다: {e}")
        return None

def get_data_as_dataframe(worksheet):
    """워크시트 데이터를 Pandas DataFrame으로 불러오고 기본 전처리를 수행합니다."""
    try:
        # 데이터가 없는 경우를 대비하여 빈 데이터프레임 생성
        data = worksheet.get_all_values()
        if not data:
            # 헤더만 있는 경우 또는 완전히 비어있는 경우
            st.warning("시트에 데이터가 없습니다. 헤더를 확인해주세요.")
            # 필수 헤더를 가진 빈 데이터프레임 반환
            headers = ['담당자', '고객명', '계약접수처', '유입경로', '날짜', '접수처월별', '전체월별', '상태']
            return pd.DataFrame(columns=headers)

        header = data[0]
        records = data[1:]
        
        # 필수 헤더 존재 여부 확인
        required_headers = ['담당자', '날짜', '계약접수처']
        for h in required_headers:
            if h not in header:
                st.error(f"시트의 첫 행에 필수 헤더 '{h}'가 없습니다. 확인해주세요.")
                return None

        df = pd.DataFrame(records, columns=header)
        df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
        # 행 번호를 추적하기 위한 인덱스 추가 (시트의 실제 행 번호와 맞춤: 헤더 1행 + 데이터 1부터 시작)
        df['row_index'] = range(2, len(df) + 2)
        return df
    except Exception as e:
        st.error(f"데이터를 DataFrame으로 변환하는 중 오류 발생: {e}")
        return None

def register_third_party_contract(worksheet, all_df):
    """타사 계약 등록 UI 및 로직을 처리합니다. (수기 입력 방식)"""
    st.header("📋 타사 계약 등록")
    st.info("계약서 파일을 업로드하고, 모든 정보를 직접 입력해주세요.")

    # --- 등록 완료 후 메일 링크 표시 로직 (재사용) ---
    if 'tp_generated_mail_url' in st.session_state:
        st.success("✅ Google Sheet에 데이터가 성공적으로 기록되었습니다.")
        st.markdown(f'<a href="{st.session_state.tp_generated_mail_url}" target="_blank" style="display: inline-block; padding: 12px 24px; background-color: #0073e6; color: white; text-decoration: none; font-weight: bold; border-radius: 5px; font-size: 16px;">📬 웍스메일 작성창 열기</a>', unsafe_allow_html=True)
        st.info("메일 작성을 완료했거나, 새 계약을 등록하려면 아래 버튼을 눌러주세요.")
        
        if st.button("🔄 새 타사 계약 등록 시작하기", use_container_width=True):
            del st.session_state.tp_generated_mail_url
            st.rerun()
        return

    # --- 파일 업로드 및 미리보기 ---
    # 1. PDF와 이미지 파일(jpg, jpeg, png)을 모두 허용
    uploaded_file = st.file_uploader(
        "계약서 파일 (PDF, JPG, PNG)을 업로드하세요.",
        type=["pdf", "jpg", "jpeg", "png"]
    )

    if uploaded_file:
        with st.expander("📄 업로드된 파일 미리보기 및 전체보기"):
            file_bytes = uploaded_file.getvalue()
            
            # 파일 타입에 따라 다른 미리보기 제공
            if uploaded_file.type == "application/pdf":
                st.markdown("##### 📄 첫 페이지 미리보기")
                # PDF는 첫 페이지를 이미지로 변환하여 표시 (기본값 0)
                preview_image = convert_pdf_page_to_image(file_bytes, page_number=0)
                if preview_image:
                    st.image(preview_image, caption="계약서 첫 페이지", use_container_width=True)
                else:
                    st.warning("PDF 미리보기를 생성할 수 없습니다.")
            else:
                # 이미지는 바로 표시
                st.image(file_bytes, caption="업로드된 이미지", use_container_width=True)

            st.markdown("---")
            st.markdown("##### 📑 전체 파일 열기/다운로드")
            st.download_button(
                label="클릭하여 전체 파일 열기",
                data=file_bytes,
                file_name=uploaded_file.name,
                mime=uploaded_file.type,
                use_container_width=True
            )

    # --- 수기 입력 폼 ---
    # 2. PDF 분석 과정 없이 모든 항목을 st.form 안에서 직접 입력
    with st.form("third_party_contract_form"):
        st.subheader("📂 계약 정보 입력")
        
        # 기본 정보
        reception_office_options = ["온라인신규", "온라인", "중고차신규", "중고차", "원큐", "노바딜", "현대캐피탈1", "현대캐피탈2", "기타"]
        inflow_channel_options = ["온라인DB", "만기", "틱톡", "홈쇼핑", "지인", "기타"]
        reception_office = st.selectbox("계약접수처", reception_office_options)
        inflow_channel = st.selectbox("유입경로", inflow_channel_options)
        col1, col2 = st.columns(2)
        with col1:
            is_additional = st.checkbox("추가")
        with col2:
            is_referral = st.checkbox("소개")
        
        st.markdown("---")

        # 계약 상세 정보 (모두 수기 입력)
        customer_name = st.text_input("고객명")
        car_model = st.text_input("대여차종")
        rental_period = st.text_input("대여기간 (개월)")
        car_price = st.text_input("차량 소비자 가격")
        monthly_fee = st.text_input("월대여료")
        deposit_prepayment = st.text_input("보증금 / 선납금")
        
        st.markdown("---")

        # 추가 정보
        commission = st.text_input("수수료")
        incentive = st.text_input("인센티브")
        delivery_date = st.text_input("투입일자")
        
        submit_button = st.form_submit_button("🚀 타사 계약 등록하기", use_container_width=True)

        if submit_button:
            # 3. 입력된 정보로 시트 저장 및 메일 생성
            user_inputs = {
                "sales_person": st.session_state['sales_person'],
                "reception_office": reception_office,
                "inflow_channel": inflow_channel
            }

            with st.spinner('Google Sheet에 데이터를 기록하는 중...'):
                try:
                    # 댓수 계산 로직 (기존과 동일)
                    current_date = datetime.now()
                    sales_person_name = user_inputs['sales_person']
                    salesperson_df = all_df[all_df['담당자'] == sales_person_name]
                    current_month_salesperson_df = salesperson_df[salesperson_df['날짜'].dt.month == current_date.month]
                    total_salesperson_monthly_count = len(current_month_salesperson_df) + 1
                    office_monthly_salesperson_df = current_month_salesperson_df[current_month_salesperson_df['계약접수처'] == reception_office]
                    total_office_salesperson_monthly_count = len(office_monthly_salesperson_df) + 1
                    
                    # 시트에 저장할 데이터 구성
                    sheet_headers = worksheet.row_values(1)
                    new_row_dict = {
                        '담당자': sales_person_name, '고객명': customer_name, '계약접수처': reception_office,
                        '유입경로': inflow_channel, '날짜': current_date.strftime("%Y-%m-%d"),
                        '접수처월별': total_office_salesperson_monthly_count, '전체월별': total_salesperson_monthly_count,
                        '상태': '정상', '추가': "O" if is_additional else "", '소개': "O" if is_referral else ""
                    }
                    new_row_list = [new_row_dict.get(h, '') for h in sheet_headers]
                    worksheet.append_row(new_row_list, value_input_option='USER_ENTERED')
                    
                    # 메일 생성을 위해 수기 입력 데이터를 딕셔너리 형태로 만듦
                    manual_data_for_mail = {
                        '고객명': customer_name, '대여차종': car_model, '대여기간': rental_period,
                        '차량 소비자 가격': car_price, '월대여료': monthly_fee, '보증금 / 선납금': deposit_prepayment
                    }
                    
                    mail_url = create_works_mail_url(
                        manual_data_for_mail, user_inputs, {
                            "office_total": total_office_salesperson_monthly_count,
                            "grand_total": total_salesperson_monthly_count
                        },
                        commission=commission, incentive=incentive,
                        delivery_date=delivery_date,
                        is_additional=is_additional, is_referral=is_referral
                    )
                    # 세션 키를 다르게 하여 기존 메뉴와 충돌 방지
                    st.session_state.tp_generated_mail_url = mail_url
                    st.rerun()

                except Exception as e:
                    st.error(f"Google Sheet 처리 중 오류 발생: {e}")

def register_novadeal_contract(worksheet, all_df):
    """노바딜 계약 등록 UI 및 로직을 처리합니다. (파일 업로드 없는 수기 입력 방식)"""
    st.header("🚗 노바딜 계약 등록")
    st.info("모든 계약 정보를 직접 입력해주세요.")

    # --- 등록 완료 후 메일 링크 표시 로직 (세션 키만 변경) ---
    if 'nd_generated_mail_url' in st.session_state:
        st.success("✅ Google Sheet에 데이터가 성공적으로 기록되었습니다.")
        st.markdown(f'<a href="{st.session_state.nd_generated_mail_url}" target="_blank" style="display: inline-block; padding: 12px 24px; background-color: #0073e6; color: white; text-decoration: none; font-weight: bold; border-radius: 5px; font-size: 16px;">📬 웍스메일 작성창 열기</a>', unsafe_allow_html=True)
        st.info("메일 작성을 완료했거나, 새 계약을 등록하려면 아래 버튼을 눌러주세요.")
        
        if st.button("🔄 새 노바딜 계약 등록 시작하기", use_container_width=True):
            del st.session_state.nd_generated_mail_url
            st.rerun()
        return

    # --- 수기 입력 폼 ---
    # 파일 업로드 및 미리보기 섹션을 완전히 제거
    with st.form("novadeal_contract_form"):
        st.subheader("📂 계약 정보 입력")
        
        # 입력 필드는 '타사 계약 등록'과 동일
        reception_office_options = ["온라인신규", "온라인", "중고차신규", "중고차", "원큐", "노바딜", "현대캐피탈1", "현대캐피탈2", "기타"]
        inflow_channel_options = ["온라인DB", "만기", "틱톡", "홈쇼핑", "지인", "기타"]
        reception_office = st.selectbox("계약접수처", reception_office_options)
        inflow_channel = st.selectbox("유입경로", inflow_channel_options)
        col1, col2 = st.columns(2)
        with col1:
            is_additional = st.checkbox("추가")
        with col2:
            is_referral = st.checkbox("소개")
        
        st.markdown("---")

        customer_name = st.text_input("고객명")
        car_model = st.text_input("대여차종")
        rental_period = st.text_input("대여기간 (개월)")
        car_price = st.text_input("차량 소비자 가격")
        monthly_fee = st.text_input("월대여료")
        deposit_prepayment = st.text_input("보증금 / 선납금")
        
        st.markdown("---")

        commission = st.text_input("수수료")
        incentive = st.text_input("인센티브")
        delivery_date = st.text_input("투입일자")
        
        submit_button = st.form_submit_button("🚀 노바딜 계약 등록하기", use_container_width=True)

        if submit_button:
            # 제출 후 로직은 '타사 계약 등록'과 동일
            user_inputs = {
                "sales_person": st.session_state['sales_person'],
                "reception_office": reception_office,
                "inflow_channel": inflow_channel
            }

            with st.spinner('Google Sheet에 데이터를 기록하는 중...'):
                try:
                    current_date = datetime.now()
                    sales_person_name = user_inputs['sales_person']
                    salesperson_df = all_df[all_df['담당자'] == sales_person_name]
                    current_month_salesperson_df = salesperson_df[salesperson_df['날짜'].dt.month == current_date.month]
                    total_salesperson_monthly_count = len(current_month_salesperson_df) + 1
                    office_monthly_salesperson_df = current_month_salesperson_df[current_month_salesperson_df['계약접수처'] == reception_office]
                    total_office_salesperson_monthly_count = len(office_monthly_salesperson_df) + 1
                    
                    sheet_headers = worksheet.row_values(1)
                    new_row_dict = {
                        '담당자': sales_person_name, '고객명': customer_name, '계약접수처': reception_office,
                        '유입경로': inflow_channel, '날짜': current_date.strftime("%Y-%m-%d"),
                        '접수처월별': total_office_salesperson_monthly_count, '전체월별': total_salesperson_monthly_count,
                        '상태': '정상', '추가': "O" if is_additional else "", '소개': "O" if is_referral else ""
                    }
                    new_row_list = [new_row_dict.get(h, '') for h in sheet_headers]
                    worksheet.append_row(new_row_list, value_input_option='USER_ENTERED')
                    
                    manual_data_for_mail = {
                        '고객명': customer_name, '대여차종': car_model, '대여기간': rental_period,
                        '차량 소비자 가격': car_price, '월대여료': monthly_fee, '보증금 / 선납금': deposit_prepayment
                    }
                    
                    mail_url = create_works_mail_url(
                        manual_data_for_mail, user_inputs, {
                            "office_total": total_office_salesperson_monthly_count,
                            "grand_total": total_salesperson_monthly_count
                        },
                        commission=commission, incentive=incentive,
                        delivery_date=delivery_date,
                        is_additional=is_additional, is_referral=is_referral
                    )
                    # 세션 키를 다르게 하여 다른 메뉴와 충돌 방지
                    st.session_state.nd_generated_mail_url = mail_url
                    st.rerun()

                except Exception as e:
                    st.error(f"Google Sheet 처리 중 오류 발생: {e}")

# --------------------------------------------------------------------------
# 2. PDF 계약서 분석 함수 (기존 코드 활용)
# --------------------------------------------------------------------------

def extract_specific_data_from_page2(pdf_file):
    """PDF 파일의 2페이지에서 지정된 데이터를 추출합니다."""
    try:
        extracted_blocks = []
        for page_layout in extract_pages(pdf_file, page_numbers=[1]):
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    extracted_blocks.append({
                        'text': element.get_text().strip(),
                        'bbox': element.bbox
                    })
        
        target_labels = {
            '고객명': ['고객명', '법인명'],
            '대여차종': ['대여차종'],
            '대여기간': ['대여기간'],
            '월대여료': ['월 대여료(VAT포함)(1)'],
            '차량 소비자 가격': ['차량 소비자 가격', '차량소비자 가격'],
            '보증금 / 선납금': ['보증금 / 선납금']
        }
        
        extracted_info = {}
        Y_TOLERANCE = 5

        for key_name, label_list in target_labels.items():
            found_value = "정보 없음"
            label_bbox = None
            for label_text in label_list:
                for block in extracted_blocks:
                    if label_text in block['text']:
                        label_bbox = block['bbox']
                        break
                if label_bbox:
                    break
            
            if label_bbox:
                label_x1, label_y0, _, label_y1 = label_bbox
                potential_values = []
                for block in extracted_blocks:
                    block_x0, block_y0, _, block_y1 = block['bbox']
                    if block_x0 > label_x1 and abs(((label_y0 + label_y1) / 2) - ((block_y0 + block_y1) / 2)) < Y_TOLERANCE:
                        potential_values.append((block_x0, block['text']))
                
                potential_values.sort(key=lambda item: item[0])
                
                if key_name == '대여기간':
                    for _, text in potential_values:
                        if text.isdigit():
                            found_value = text
                            break
                elif key_name in ['월대여료', '차량 소비자 가격']:
                    for _, text in potential_values:
                        match = re.search(r'[\d,]+', text)
                        if match:
                            found_value = match.group(0)
                            break
                elif key_name == '보증금 / 선납금':
                    money_values = []
                    for _, text in potential_values:
                        matches = re.findall(r'\d{1,3}(?:,\d{3})*|\d+', text)
                        money_values.extend(matches)
                    if len(money_values) >= 2:
                        found_value = f"보증금: {money_values[0]} / 선납금: {money_values[1]}"
                    elif len(money_values) == 1:
                        found_value = f"보증금/선납금: {money_values[0]}"
                elif key_name == '대여차종':
                    if potential_values:
                        full_model_name = potential_values[0][1]
                        found_value = summarize_car_model(full_model_name)
                else:
                    if potential_values:
                        found_value = potential_values[0][1]

            extracted_info[key_name] = found_value
        return extracted_info
    except Exception as e:
        return {"오류": str(e)}

def summarize_car_model(full_model_name):
    """차량 모델명을 간소화합니다."""
    stop_patterns = [r'\d\.\d', r'\d{2}"', '2WD', '4WD', 'AWD', r'\sAT', r'\sMT', r'\/', '디젤', '가솔린', 'LPi', 'LPG', '하이브리드', '터보', '기본']
    first_cut_index = len(full_model_name)
    for pattern in stop_patterns:
        match = re.search(pattern, full_model_name)
        if match and match.start() < first_cut_index:
            first_cut_index = match.start()
    return full_model_name[:first_cut_index].strip()

# --------------------------------------------------------------------------
# 3. UI 렌더링 함수
# --------------------------------------------------------------------------

def show_login_screen():
    """로그인 화면 UI를 표시합니다."""
    st.title("📄 계약 처리 자동화 시스템")
    st.subheader("담당자 이름을 입력해주세요.")
    
    sales_person = st.text_input("담당자 이름", key="login_name_input")
    
    if st.button("로그인", key="login_button"):
        if sales_person:
            st.session_state['logged_in'] = True
            st.session_state['sales_person'] = sales_person
            st.rerun()
        else:
            st.warning("담당자 이름을 입력해야 합니다.")

def show_main_app():
    """메인 애플리케이션 화면 UI를 표시합니다."""
    st.sidebar.header(f"👤 {st.session_state['sales_person']}님")
    
    # 1. 사이드바 메뉴를 새로운 상위 메뉴 구조로 변경
    mode = st.sidebar.radio(
        "원하는 작업을 선택하세요.",
        ('내 계약 조회', '계약 등록', '계약 수정', '계약 취소') # ◀️ 메뉴 단순화
    )
    
    worksheet = connect_to_sheet()
    if worksheet is None: return

    df = get_data_as_dataframe(worksheet)
    if df is None: return

    user_df = df[(df['담당자'] == st.session_state['sales_person']) & (df['상태'] != '취소')]

    # 2. '계약 등록' 메뉴 선택 시, 새로 만든 서브메뉴 함수를 호출
    if mode == '내 계약 조회':
        view_contracts(user_df)
    elif mode == '계약 등록':
        show_registration_submenu(worksheet, df) # ◀️ 서브메뉴 함수 호출
    elif mode == '계약 수정':
        edit_contract(worksheet, user_df)
    elif mode == '계약 취소':
        cancel_contract(worksheet, user_df)

    if st.sidebar.button("로그아웃"):
        st.session_state['logged_in'] = False
        st.rerun()

def view_contracts(user_df):
    """담당자의 계약 목록을 표시합니다."""
    st.header("나의 계약 목록")
    if user_df.empty:
        st.info("등록된 계약이 없습니다.")
    else:
        # 화면에 표시할 컬럼만 선택
        display_cols = ['날짜', '고객명', '계약접수처', '유입경로', '상태']
        # user_df에 있는 컬럼만 필터링
        display_cols = [col for col in display_cols if col in user_df.columns]
        
        # '날짜' 컬럼을 문자열로 변환하여 시간 정보 제거
        df_display = user_df.copy()
        df_display['날짜'] = df_display['날짜'].dt.strftime('%Y-%m-%d')
        
        st.dataframe(df_display[display_cols], use_container_width=True)


def show_registration_submenu(worksheet, df):
    """'계약 등록' 선택 시, 세부 등록 유형을 탭으로 보여주는 함수"""
    st.header("📑 계약 등록")
    st.info("등록할 계약 유형을 선택하세요.")

    # st.tabs를 사용하여 세 가지 등록 메뉴를 생성
    tab_lotte, tab_third_party, tab_novadeal = st.tabs([
        "롯데 계약 (자동 분석)", 
        "타사 계약 (수기 입력)", 
        "노바딜 계약 (수기 입력)"
    ])

    # 각 탭(Tab) 내부를 정의
    with tab_lotte:
        # '롯데 계약' 탭을 클릭하면 register_lotte_contract 함수가 실행됨
        register_lotte_contract(worksheet, df)

    with tab_third_party:
        # '타사 계약' 탭을 클릭하면 register_third_party_contract 함수가 실행됨
        register_third_party_contract(worksheet, df)

    with tab_novadeal:
        # '노바딜 계약' 탭을 클릭하면 register_novadeal_contract 함수가 실행됨
        register_novadeal_contract(worksheet, df)

def register_lotte_contract(worksheet, all_df):
    """신규 계약 등록 UI 및 로직을 처리합니다. (입력폼 통합 버전)"""
    st.header("신규 계약 등록")

    # (UI Part 1: 등록 완료 후 메일 링크 표시 로직은 동일)
    if 'generated_mail_url' in st.session_state:
        st.success("✅ Google Sheet에 데이터가 성공적으로 기록되었습니다.")
        st.markdown(f'<a href="{st.session_state.generated_mail_url}" target="_blank" style="display: inline-block; padding: 12px 24px; background-color: #0073e6; color: white; text-decoration: none; font-weight: bold; border-radius: 5px; font-size: 16px;">📬 웍스메일 작성창 열기</a>', unsafe_allow_html=True)
        st.info("메일 작성을 완료했거나, 새 계약을 등록하려면 아래 버튼을 눌러주세요.")
        
        if st.button("🔄 새 계약 등록 시작하기", use_container_width=True):
            del st.session_state.generated_mail_url
            st.rerun()
        return

    # --- UI Part 2: 계약 등록 폼 ---
    # [변경] 파일 업로더만 남기고 다른 위젯들은 st.form 안으로 이동
    uploaded_file = st.file_uploader("계약서 PDF 파일을 업로드하세요.", type="pdf")

    # (PDF 분석 로직은 동일)
    if uploaded_file is not None:
        if 'last_uploaded_filename' not in st.session_state or st.session_state.last_uploaded_filename != uploaded_file.name:
            with st.spinner('계약서를 분석 중...'):
                st.session_state.extracted_data = extract_specific_data_from_page2(uploaded_file)
                st.session_state.last_uploaded_filename = uploaded_file.name
                st.success("✅ 계약서 정보 추출 완료!")
    
    # (미리보기 로직은 동일)
    if uploaded_file: # 파일이 업로드된 상태라면 미리보기 섹션 표시
        with st.expander("📄 업로드된 계약서 미리보기 및 전체보기"):
            pdf_bytes = uploaded_file.getvalue()
            st.markdown("##### 📄 두 번째 페이지 미리보기")
            preview_image = convert_pdf_page_to_image(pdf_bytes)
            if preview_image:
                st.image(preview_image, caption="계약서 두 번째 페이지", use_container_width=True)
            else:
                st.warning("미리보기를 생성할 수 없습니다.")
            st.markdown("---")
            st.markdown("##### 📑 전체 파일 열기")
            st.download_button(
                label="클릭하여 전체 계약서 열기",
                data=pdf_bytes,
                file_name=uploaded_file.name,
                mime="application/pdf",
                use_container_width=True
            )

    if 'extracted_data' in st.session_state and st.session_state.extracted_data:
        if "오류" in st.session_state.extracted_data:
            st.error(f"PDF 분석 오류: {st.session_state.extracted_data['오류']}")
            del st.session_state.extracted_data
            return

        with st.form("edit_and_submit_form"):
            # [변경] 모든 입력 위젯을 form의 맨 위로 통합
            st.subheader("📂 계약 기본 정보")
            reception_office_options = ["온라인신규", "온라인", "중고차신규", "중고차", "원큐", "노바딜", "현대캐피탈1", "현대캐피탈2", "기타"]
            inflow_channel_options = ["온라인DB", "만기", "틱톡", "홈쇼핑", "지인", "기타"]
            
            reception_office = st.selectbox("계약접수처", reception_office_options)
            inflow_channel = st.selectbox("유입경로", inflow_channel_options)

            col1, col2 = st.columns(2)
            with col1:
                is_additional = st.checkbox("추가")
            with col2:
                is_referral = st.checkbox("소개")

            st.markdown("---")
            st.subheader("📝 추출된 계약 정보 (수정 가능)")
            edited_data = {}
            for key, value in st.session_state.extracted_data.items():
                display_value = "" if value == "정보 없음" else value
                edited_data[key] = st.text_input(f"**{key}**", value=display_value)
            
            st.markdown("---")
            st.subheader("💰 추가 정보")
            commission_input = st.text_input("수수료")
            incentive_input = st.text_input("인센티브")
            delivery_date_input = st.text_input("투입일자")

            submit_button = st.form_submit_button("🚀 시트에 최종 등록하기", use_container_width=True)

            if submit_button:
                # (이하 제출 로직은 모두 동일)
                user_inputs = { "sales_person": st.session_state['sales_person'], "reception_office": reception_office, "inflow_channel": inflow_channel }
                
                with st.spinner('Google Sheet에 데이터를 기록하는 중...'):
                    try:
                        current_date = datetime.now()
                        sales_person_name = user_inputs['sales_person']
                        salesperson_df = all_df[all_df['담당자'] == sales_person_name]
                        current_month_salesperson_df = salesperson_df[salesperson_df['날짜'].dt.month == current_date.month]
                        total_salesperson_monthly_count = len(current_month_salesperson_df) + 1
                        office_monthly_salesperson_df = current_month_salesperson_df[current_month_salesperson_df['계약접수처'] == reception_office]
                        total_office_salesperson_monthly_count = len(office_monthly_salesperson_df) + 1
                        
                        sheet_headers = worksheet.row_values(1)
                        new_row_dict = {
                            '담당자': sales_person_name,
                            '고객명': edited_data.get('고객명', 'N/A'),
                            '계약접수처': user_inputs['reception_office'],
                            '유입경로': user_inputs['inflow_channel'],
                            '날짜': current_date.strftime("%Y-%m-%d"),
                            '접수처월별': total_office_salesperson_monthly_count,
                            '전체월별': total_salesperson_monthly_count,
                            '상태': '정상',
                            '추가': "O" if is_additional else "",
                            '소개': "O" if is_referral else ""
                        }
                        
                        new_row_list = [new_row_dict.get(h, '') for h in sheet_headers]
                        worksheet.append_row(new_row_list, value_input_option='USER_ENTERED')
                        
                        mail_url = create_works_mail_url(
                            edited_data, user_inputs, {
                                "office_total": total_office_salesperson_monthly_count,
                                "grand_total": total_salesperson_monthly_count
                            },
                            commission=commission_input, incentive=incentive_input,
                            delivery_date=delivery_date_input,
                            is_additional=is_additional, is_referral=is_referral
                        )
                        st.session_state.generated_mail_url = mail_url

                        del st.session_state.extracted_data
                        del st.session_state.last_uploaded_filename
                        st.rerun()

                    except Exception as e:
                        st.error(f"Google Sheet 처리 중 오류 발생: {e}")

def edit_contract(worksheet, user_df):
    """계약 수정 UI 및 로직을 처리합니다."""
    st.header("계약 수정")
    if user_df.empty:
        st.info("수정할 계약이 없습니다.")
        return

    # 선택을 위한 고유 식별자 생성
    user_df['display'] = user_df.apply(lambda row: f"{row['날짜'].strftime('%Y-%m-%d')} / {row['고객명']}", axis=1)
    
    selected_contract_display = st.selectbox(
        "수정할 계약을 선택하세요.",
        user_df['display'],
        index=None,
        placeholder="계약 선택..."
    )

    if selected_contract_display:
        selected_row = user_df[user_df['display'] == selected_contract_display].iloc[0]
        row_to_edit_index = selected_row['row_index']

        with st.form("edit_form"):
            st.write(f"**고객명:** {selected_row['고객명']}")
            
            # 수정 가능한 필드들
            new_reception_office = st.text_input("계약접수처", value=selected_row.get('계약접수처', ''))
            new_inflow_channel = st.text_input("유입경로", value=selected_row.get('유입경로', ''))
            
            submitted = st.form_submit_button("수정 내용 저장")
            if submitted:
                try:
                    # gspread는 1-based index를 사용합니다.
                    # B, C, D... 열에 해당. A열(담당자)은 1, B열은 2...
                    # 헤더를 기준으로 열 인덱스를 동적으로 찾기
                    headers = worksheet.row_values(1)
                    office_col = headers.index('계약접수처') + 1
                    inflow_col = headers.index('유입경로') + 1

                    worksheet.update_cell(row_to_edit_index, office_col, new_reception_office)
                    worksheet.update_cell(row_to_edit_index, inflow_col, new_inflow_channel)
                    
                    st.success("계약 정보가 성공적으로 수정되었습니다.")
                    st.info("페이지가 곧 새로고침됩니다.")
                    st.rerun() # 수정 후 화면을 새로고침하여 최신 상태를 반영
                except Exception as e:
                    st.error(f"수정 중 오류가 발생했습니다: {e}")

def cancel_contract(worksheet, user_df):
    """계약 취소 UI 및 로직을 처리합니다."""
    st.header("계약 취소")
    if user_df.empty:
        st.info("취소할 계약이 없습니다.")
        return

    user_df['display'] = user_df.apply(lambda row: f"{row['날짜'].strftime('%Y-%m-%d')} / {row['고객명']}", axis=1)
    
    selected_contract_display = st.selectbox(
        "취소할 계약을 선택하세요.",
        user_df['display'],
        index=None,
        placeholder="계약 선택..."
    )

    if selected_contract_display:
        st.warning(f"**'{selected_contract_display}'** 계약을 정말 취소하시겠습니까? 이 작업은 되돌릴 수 없습니다.")
        
        if st.button("🔴 예, 계약을 취소합니다.", use_container_width=True):
            selected_row = user_df[user_df['display'] == selected_contract_display].iloc[0]
            row_to_cancel_index = selected_row['row_index']
            
            try:
                headers = worksheet.row_values(1)
                if '상태' not in headers:
                    st.error("시트에 '상태' 컬럼이 없습니다. '상태' 컬럼을 추가해주세요.")
                    return
                
                status_col = headers.index('상태') + 1
                worksheet.update_cell(row_to_cancel_index, status_col, "취소")
                st.success("계약이 성공적으로 취소 처리되었습니다.")
                st.info("페이지가 곧 새로고침됩니다.")
                st.rerun()
            except Exception as e:
                st.error(f"취소 처리 중 오류가 발생했습니다: {e}")

def create_works_mail_url(extracted_data, user_inputs, calculated_totals, commission, incentive, delivery_date, is_additional, is_referral):
    """Naver Works Mail 작성 URL을 생성합니다. (투입일자 추가 버전)"""
    # (기존 변수 선언은 동일)
    customer_name = extracted_data.get('고객명', '')
    car_model = extracted_data.get('대여차종', '')
    rental_period = extracted_data.get('대여기간', '')
    car_price = extracted_data.get('차량 소비자 가격', '')
    monthly_fee = extracted_data.get('월대여료', '')
    deposit_prepayment = extracted_data.get('보증금 / 선납금', '')
    
    sales_person = user_inputs.get('sales_person', '')
    reception_office = user_inputs.get('reception_office', '')
    inflow_channel = user_inputs.get('inflow_channel', '')
    
    office_total = calculated_totals.get('office_total', 0)
    grand_total = calculated_totals.get('grand_total', 0)

    # (subject 생성 로직은 동일)
    status_text = []
    if is_additional:
        status_text.append("추가")
    if is_referral:
        status_text.append("소개")
    
    final_status = f" / { ' / '.join(status_text) }" if status_text else ""

    subject = f"{sales_person} / {customer_name} / {reception_office} / {office_total} / {inflow_channel} / {grand_total}{final_status}"
    
    # [변경] body의 '투입일자' 항목에 전달받은 delivery_date 값을 채워 넣습니다.
    body = f"""고객명 : {customer_name}
대여차종 : {car_model}
수수료 : {commission}
대여기간 : {rental_period}
차량 소비자 가격 : {car_price}
월대여료 : {monthly_fee}
보증금/선납금 : {deposit_prepayment}
투입일자 : {delivery_date}
인센티브 : {incentive}"""

    # (URL 생성 로직은 동일)
    base_url = "https://mail.worksmobile.com/write/popup"
    to_param = "문정동사서함 <automedia@automediarentcar.com>"
    
    final_url = (f"{base_url}?to={urllib.parse.quote(to_param)}"
                 f"&subject={urllib.parse.quote(subject)}"
                 f"&body={urllib.parse.quote(body)}"
                 f"&orderType=new&memo=false")
    return final_url



# --------------------------------------------------------------------------
# 4. Streamlit 앱 실행 로직
# --------------------------------------------------------------------------
st.set_page_config(page_title="계약 처리 자동화", layout="centered")

# 세션 상태 초기화
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# 로그인 상태에 따라 다른 화면 표시
if st.session_state['logged_in']:
    show_main_app()
else:
    show_login_screen()
