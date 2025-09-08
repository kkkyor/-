import streamlit as st
import re
from io import StringIO
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
import pandas as pd
import gspread
import urllib.parse
from datetime import datetime
from gspread_dataframe import set_with_dataframe
import urllib.parse

# --------------------------------------------------------------------------
# 1. 보내주신 PDF 추출 함수 (수정 없이 거의 그대로 사용)
# --------------------------------------------------------------------------
def extract_specific_data_from_page2(pdf_path):
    """
    pdfminer.six를 사용하여 PDF 2페이지의 레이아웃을 분석하고,
    특정 항목들의 값을 추출하여 딕셔너리로 반환합니다.
    """
    try:
        # 1. PDF 2페이지만 타겟으로 레이아웃 객체 추출
        extracted_blocks = []
        # page_numbers=[1]은 0-based 인덱싱으로 2페이지를 의미
        for page_layout in extract_pages(pdf_path, page_numbers=[1]):
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    extracted_blocks.append({
                        'text': element.get_text().strip(),
                        'bbox': element.bbox  # (x0, y0, x1, y1)
                    })

        # 2. 추출할 목표 레이블 정의 (리스트 형태로 다중 레이블 지정 가능)
        target_labels = {
            '고객명': ['고객명', '법인명'],
            '대여차종': ['대여차종'],
            '대여기간': ['대여기간'],
            '월대여료': ['월 대여료(VAT포함)(1)'],
            # '차량 소비자 가격'과 '차량소비자 가격' 둘 다 찾도록 수정
            '차량 소비자 가격': ['차량 소비자 가격', '차량소비자 가격'],
            '보증금 / 선납금': ['보증금 / 선납금']
        }
        
        extracted_info = {}
        Y_TOLERANCE = 5  # 같은 라인에 있는지 판단하기 위한 Y좌표 허용 오차

        # 3. 레이블을 기준으로 값 찾기
        for key_name, label_list in target_labels.items():
            found_value = "정보를 찾을 수 없음"
            label_bbox = None

            # 레이블 텍스트 목록을 순회하며 일치하는 블록 찾기
            for label_text in label_list:
                for block in extracted_blocks:
                    if label_text in block['text']:
                        label_bbox = block['bbox']
                        break  # 블록 루프 탈출
                if label_bbox:
                    break  # 레이블 목록 루프 탈출
            
            if label_bbox:
                label_x1, label_y0, _, label_y1 = label_bbox
                potential_values = []
                
                # 레이블의 오른쪽에 있고, Y좌표가 비슷한 블록들을 후보로 추가
                for block in extracted_blocks:
                    block_x0, block_y0, _, block_y1 = block['bbox']
                    # y좌표의 중간값이 허용 오차 이내에 있고, x좌표가 레이블의 오른쪽에 있는지 확인
                    if block_x0 > label_x1 and abs(((label_y0 + label_y1) / 2) - ((block_y0 + block_y1) / 2)) < Y_TOLERANCE:
                        potential_values.append((block_x0, block['text']))
                
                # X좌표 기준으로 정렬하여 가장 가까운 값을 찾음
                potential_values.sort(key=lambda item: item[0])
                
                # 키에 따라 값 추출 로직 분기
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
                        # 숫자와 콤마로 이루어진 값만 찾음
                        matches = re.findall(r'\d{1,3}(?:,\d{3})*|\d+', text)
                        money_values.extend(matches)
                    
                    if len(money_values) >= 2:
                        found_value = f"보증금: {money_values[0]} / 선납금: {money_values[1]}"
                    elif len(money_values) == 1:
                        found_value = f"보증금/선납금: {money_values[0]}"
                
                elif key_name == '대여차종':
                    if potential_values:
                        full_model_name = potential_values[0][1]
                        # 요약 함수를 호출하여 모델명 단축
                        found_value = summarize_car_model(full_model_name)

                else: # 고객명 등 나머지 항목
                    if potential_values:
                        found_value = potential_values[0][1]

            extracted_info[key_name] = found_value
            
        return extracted_info

    except Exception as e:
        return {"오류": str(e)}

def summarize_car_model(full_model_name):
    """
    복잡한 차량 모델명에서 핵심적인 모델명만 추출하여 요약합니다.
    예: "G80 (G)2.5T 18"/기본 2WD AT" -> "G80 (G)"
    """
    # 상세 스펙을 나타내는 패턴 목록
    # 이 목록에 있는 키워드나 패턴이 처음 나타나는 위치에서 문자열을 자릅니다.
    stop_patterns = [
        r'\d\.\d',          # 엔진 배기량 (예: 2.5, 3.8)
        r'\d{2}"',          # 휠 사이즈 (예: 18")
        r'2WD', '4WD', 'AWD', # 구동 방식
        r'\sAT', r'\sMT',   # 변속기
        r'\/',              # 슬래시 구분자
        '디젤', '가솔린', 'LPi', 'LPG', '하이브리드', '터보', # 연료/엔진 타입
        '기본'              # '기본' 트림
    ]

    first_cut_index = len(full_model_name)

    # 각 패턴이 가장 먼저 나타나는 위치를 찾음
    for pattern in stop_patterns:
        # re.search는 문자열 전체에서 패턴을 검색합니다.
        match = re.search(pattern, full_model_name)
        if match and match.start() < first_cut_index:
            first_cut_index = match.start()

    # 가장 먼저 나타난 스펙 패턴 이전까지의 문자열을 공백 제거 후 반환
    return full_model_name[:first_cut_index].strip()


# app.py에 있는 기존 update_spreadsheet_and_calculate_totals 함수를
# 아래의 새로운 함수로 완전히 대체합니다.

def update_spreadsheet_and_calculate_totals(extracted_data, user_inputs):
    """
    (개선된 버전) Google Sheet에 연결하여 월간 합산을 계산하고 새 데이터를 기록합니다.
    시트가 비어있거나 헤더가 잘못된 경우를 방어합니다.
    """
    try:
        # 1. Google Sheets API 인증
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        spreadsheet = gc.open("테스트0905")
        worksheet = spreadsheet.sheet1

        # 2. 헤더와 모든 데이터를 별도로 불러오기
        header = worksheet.row_values(1)
        records = worksheet.get_all_values()[1:] # 헤더를 제외한 실제 데이터

        # ✨ 방어 코드 1: 필수 헤더가 존재하는지 먼저 확인
        required_headers = ['날짜', '계약접수처']
        for h in required_headers:
            if h not in header:
                return {"status": "error", "message": f"시트의 1행에 '{h}' 헤더가 없습니다. 헤더를 확인해주세요."}

        # 3. Pandas DataFrame으로 변환 (헤더를 명시적으로 지정)
        df = pd.DataFrame(records, columns=header)

        # 4. 월간 합산 계산
        current_date = datetime.now()
        current_month = current_date.month

        if not df.empty:
            df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
            monthly_df = df[df['날짜'].dt.month == current_month]
        else:
            monthly_df = pd.DataFrame()

        # ✨ 방어 코드 2: 계산 시에도 헤더 존재 여부 재확인
        total_monthly_count = len(monthly_df) + 1
        
        if '계약접수처' in monthly_df.columns:
            office_monthly_df = monthly_df[monthly_df['계약접수처'] == user_inputs['reception_office']]
            total_office_monthly_count = len(office_monthly_df) + 1
        else:
            # 헤더가 있지만 데이터가 없는 초기 상태일 경우
            total_office_monthly_count = 1
        
        # 5. 시트에 추가할 새로운 행 데이터 생성 (헤더 순서에 맞게 리스트로)
        new_row_list = [
            user_inputs['sales_person'],
            extracted_data.get('고객명', 'N/A'),
            user_inputs['reception_office'],
            user_inputs['inflow_channel'],
            current_date.strftime("%Y-%m-%d"),
            total_office_monthly_count,
            total_monthly_count
        ]

        # 6. 새로운 행을 시트에 추가
        worksheet.append_row(new_row_list, value_input_option='USER_ENTERED')
        
        return {
            "status": "success",
            "message": "Google Sheet에 데이터를 성공적으로 기록했습니다.",
            "office_total": total_office_monthly_count,
            "grand_total": total_monthly_count
}

    except Exception as e:
        # 더 구체적인 오류 메시지를 반환하도록 수정
        return {"status": "error", "message": f"오류 발생: {type(e).__name__} - {str(e)}"}

def create_works_mail_url(extracted_data, user_inputs, calculated_totals):
    """
    추출된 데이터, 사용자 입력, 계산된 합산 값을 조합하여
    Naver Works Mail 작성 URL을 생성합니다.
    """
    # URL에 들어갈 값들을 변수로 정리
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

    # 1. 이메일 제목(Subject) 생성
    subject = f"{sales_person} / {customer_name} / {reception_office} / {office_total} / {inflow_channel} / {grand_total}"

    # 2. 이메일 본문(Body) 생성 (\n은 줄바꿈)
    body = f"""고객명 : {customer_name}
대여차종 : {car_model}
수수료 : 
대여기간 : {rental_period}
차량 소비자 가격 : {car_price}
월대여료 : {monthly_fee}
보증금/선납금 : {deposit_prepayment}
투입일자 : 
인센티브 : """

    # 3. Base URL 및 고정 수신자 정보
    base_url = "https://mail.worksmobile.com/write/popup"
    # 받는 사람 정보는 URL 인코딩이 필요할 수 있으므로 미리 변수로 지정
    to_param = "문정동사서함 <automedia@automediarentcar.com>"

    # 4. 최종 URL 조립 (한글 등 특수문자가 깨지지 않도록 URL 인코딩 필수!)
    final_url = (
        f"{base_url}"
        f"?to={urllib.parse.quote(to_param)}"
        f"&subject={urllib.parse.quote(subject)}"
        f"&body={urllib.parse.quote(body)}"
        f"&orderType=new&memo=false"
    )
    
    return final_url

# --------------------------------------------------------------------------
# 2. Streamlit 웹 UI 구성
# --------------------------------------------------------------------------
st.set_page_config(page_title="계약 처리 자동화", layout="centered")
st.title("📄 계약서 처리 및 메일 자동화")

# 사이드바에 사용자 입력 필드 배치
st.sidebar.header("📝 정보 입력")
sales_person = st.sidebar.text_input("담당자 이름")
reception_office = st.sidebar.text_input("계약접수처")
inflow_channel = st.sidebar.text_input("유입경로")

# 메인 화면에 파일 업로더 배치
uploaded_file = st.file_uploader("계약서 PDF 파일을 업로드하세요.", type="pdf")

process_button = st.button("🚀 처리 및 메일 링크 생성", use_container_width=True)

# --------------------------------------------------------------------------
# 3. 버튼 클릭 시 모든 로직 실행
# --------------------------------------------------------------------------
if process_button:
    # 입력 값 검증
    if not all([sales_person, reception_office, inflow_channel, uploaded_file]):
        st.warning("모든 정보를 입력하고 파일을 업로드해주세요.")
    else:
        # 사용자 입력을 딕셔너리로 묶기
        user_inputs = {
            "sales_person": sales_person,
            "reception_office": reception_office,
            "inflow_channel": inflow_channel
        }

        with st.spinner('계약서를 분석하고 있습니다...'):
            extracted_data = extract_specific_data_from_page2(uploaded_file)
        
        if "오류" in extracted_data:
            st.error(f"PDF 분석 중 오류가 발생했습니다: {extracted_data['오류']}")
        else:
            st.success("✅ 계약서 정보 추출 완료!")
            st.write(extracted_data)

            # --- ✨ Google Sheet 연동 함수 호출! ---
            with st.spinner('Google Sheet에 데이터를 기록하고 합산을 계산하는 중...'):
                sheet_result = update_spreadsheet_and_calculate_totals(extracted_data, user_inputs)
            
            if sheet_result['status'] == 'success':
                st.success(f"✅ {sheet_result['message']}")
                
                # ✨ --- URL 생성 및 링크 표시 로직 추가 ---
                
                # 1. 시트 함수에서 반환된 합산 값 저장
                calculated_totals = {
                    "office_total": sheet_result['office_total'],
                    "grand_total": sheet_result['grand_total']
                }
                
                # 2. URL 생성 함수 호출
                mail_url = create_works_mail_url(extracted_data, user_inputs, calculated_totals)
                
                # 3. 클릭 가능한 링크(새 창)로 화면에 표시
                st.markdown(f'''
                <a href="{mail_url}" target="_blank" style="
                    display: inline-block;
                    padding: 10px 20px;
                    background-color: #0073e6;
                    color: white;
                    text-decoration: none;
                    font-weight: bold;
                    border-radius: 5px;">
                    📬 웍스메일 작성창 열기
                </a>
                ''', unsafe_allow_html=True)

            else:
                st.error(f"❗️ Google Sheet 처리 중 오류 발생:\n{sheet_result['message']}")