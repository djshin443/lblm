import requests
import json
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re  # 정규 표현식을 사용하기 위해 re 모듈을 임포트합니다.
import os
import time
import sys

# 사용자 정의 상태 변수 확장
# 사용자 정의 상태 변수 확장
ON_WEBSITE_SOLD = "O"  # 공홈에도 있고 실제로 판매 중인 제품
NOT_ON_WEBSITE_SOLD = "X"  # 공홈에는 없지만 실제로 판매 중인 제품
NOT_AVAILABLE = "N/A"  # 공홈에도 없고 실제로 판매하지 않는 제품
ON_WEBSITE_NOT_SOLD = "O"  # 공홈에는 있지만 판매하지 않는 제품
max_stores_to_save = 5  # 0이면 매장 전체 출력)

current_dir = os.getcwd()
datetime_now = datetime.now().strftime('%Y%m%d')

source_filename = '루이비통.xlsx'
target_filename = f'루이비통_{datetime_now}.xlsx'
progress_file_name = f'progress_{datetime_now}.txt'  # 현재 날짜를 포함한 파일 이름

source_path = os.path.join(current_dir, source_filename)
target_path = os.path.join(current_dir, target_filename)
progress_file = os.path.join(current_dir, progress_file_name)

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))

def save_progress(progress_file, sheet_name, row):
    with open(progress_file, 'w') as file:
        file.write(f"{sheet_name},{row}")

def load_progress(progress_file):
    if os.path.exists(progress_file):
        with open(progress_file, 'r') as file:
            content = file.read()
            if content:
                sheet_name, row = content.split(',')
                return sheet_name, int(row)
    return None, 0

# Selenium 설정
options = Options()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument('--disable-infobars')
options.add_argument('--headless')  # 필요한 경우 주석 해제
options.add_argument('--window-size=1920,1080')
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
options.add_argument('--log-level=3')  # 로그 레벨 추가

# 웹 드라이버 초기화
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.implicitly_wait(30)

try:
    # 엑셀 파일 불러오기 및 복사
    workbook = load_workbook(filename=source_path)
    workbook.save(filename=target_path)
    print(f"원본 파일을 '{target_filename}'으로 복사했습니다.")
except Exception as e:
    print(f"파일 복사 중 오류: {e}")
    sys.exit()

# 진행 상황 로드
last_processed_sheet, last_processed_row = load_progress(progress_file)

# 제품 정보 및 매장 가격 가져오기 함수
def get_product_info(product_code):
    try:
        # 웹 페이지에서 제품 정보 가져오기
        driver.get(f'https://kr.louisvuitton.com/kor-kr/search/{product_code}')
        print("현재 페이지 URL:", driver.current_url)

        # 가격 정보 추출
        try:
            price_element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'lv-product__price'))
            )
            price_text = price_element.text.strip()
            # '₩' 기호와 추가 공백을 제거합니다.
            price_text = re.sub(r'[₩\s]', '', price_text)
        except Exception as e:
            print("가격 정보를 불러오는 중 예외 발생:", e)
            price_text = "N/A"

        # 제품 이름 추출
        try:
            product_name = driver.find_element(By.CLASS_NAME, 'lv-product__name').text
        except:
            product_name = "N/A"

        # 소재 정보 추출
        try:
            material = driver.find_element(By.CLASS_NAME, 'lv-product-variation-selector__value').text
        except:
            material = "N/A"

        # 제품 세부 정보 추출
        try:
            details_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '.lv-product-features-buttons__button'))
            )
            details_button.click()
            description_text = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, '#modalContent p'))
            ).text
        except:
            description_text = "N/A"

        # 사이즈 정보 추출
        try:
            size_element = driver.find_element(By.CSS_SELECTOR, '#modalContent > div.lv-product-detailed-features > div > div.lv-product-dimension.body-s.lv-product-detailed-features__dimensions')
            if size_element:
                size_text = size_element.text.strip()  
                size_text = ' '.join(size_text.splitlines())  
            else:
                size_text = "N/A"

        except Exception as e:
            size_text = "N/A"

        # 세부 특징 추출
        try:
            features = driver.find_elements(By.CSS_SELECTOR, '.lv-product-detailed-features__description li')
            feature_text = "\n".join([feature.text.strip() for feature in features if feature.text.strip() != ""])
        except Exception as e:
            feature_text = "N/A"

        return price_text, product_name, material, size_text, feature_text, description_text

    except Exception as e:
        print("An error occurred:", e)
        return "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"

# 각 제품 코드에 대한 정보 가져오기
for sheet_name in workbook.sheetnames:
    if last_processed_sheet and sheet_name < last_processed_sheet:
        continue
    sheet = workbook[sheet_name]
    print(f"시트 '{sheet_name}' 처리 중...")

    for row_num, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if last_processed_sheet == sheet_name and row_num <= last_processed_row:
            continue

        product_code = row[0]
        product_price, product_name, material, size, features, description = get_product_info(product_code)
        print(f"제품 코드: {product_code}, 제품명: {product_name}, 가격: {product_price}, 소재: {material}, 사이즈: {size}")
        print(f"특징: {features}, 설명: {description}")

        # 매장 재고 조회 성공 후 처리
        total_stores_count = 0
        seoul_dosan_count = 0  # 서울/도산 매장 수 초기화
        stock_info = []  # 재고가 있는 매장 이름을 저장할 리스트

        if product_price != 'N/A':  # 가격 정보가 'N/A'가 아닌 경우에만 재고 조회를 실행
            try:
                # 매장 재고 조회
                total_stores_count = 0
                seoul_dosan_count = 0  # 서울/도산 매장 수 초기화
                stock_info = []  # 재고가 있는 매장 이름을 저장할 리스트

                if product_price != 'N/A':  # 가격 정보가 'N/A'가 아닌 경우에만 재고 조회를 실행
                    try:
                        url = 'https://api.louisvuitton.com/eco-eu/search-merch-eapi/v1/kor-kr/stores/query'
                        headers = {
                                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
                                'Content-Type': 'application/json',
                                'Accept': 'application/json, text/plain, */*',
                                'client_secret': '60bbcdcD722D411B88cBb72C8246a22F',
                                'client_id': '607e3016889f431fb8020693311016c9',
                                'Origin': 'https://kr.louisvuitton.com',
                                'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
                                'Cookie': '123'
                            }
                        data = {
                                "flagShip": False,
                                "country": "KR",
                                "query": "서울",
                                "latitudeA": "37.768825744921735",
                                "latitudeB": "37.36061764438869",
                                "latitudeCenter": "37.56472169465521",
                                "longitudeA": "126.51683339179687",
                                "longitudeB": "127.43144520820312",
                                "longitudeCenter": "126.97413929999999",
                                "query": "",
                                "clickAndCollect": False,
                                "skuId": product_code,  
                                "pageType": "productsheet"
                            }
                        response = requests.post(url, headers=headers, data=json.dumps(data))
                        #print(f"API 응답: {response.text}")  # 로깅
                        
                        if response.status_code == 200:
                            response_data = response.json()
                            stores_with_stock = [
                                store for store in response_data.get('hits', [])
                                if any(
                                    prop.get('value') == 'true' and prop.get('name') == 'stockAvailability'
                                    for prop in store.get('additionalProperty', [])
                                )
                            ]
                            total_stores_count = len(stores_with_stock)

                            for store in stores_with_stock:
                                store_name = store.get('name')
                                stock_info.append(store_name)

                                # 서울/도산 매장 카운트
                                if '서울' in store_name or '도산' in store_name:
                                    seoul_dosan_count += 1

                        else:
                            print(f"API 요청이 실패했습니다. 상태 코드: {response.status_code}")

                    except Exception as e:
                        print(f"API 요청 중 오류 발생: {e}")
                        continue

                # 재고 정보 출력
                print(f"총 {total_stores_count} 개 매장에서 재고를 찾았습니다.")
                print(f"서울/도산 매장 중 {seoul_dosan_count} 개 매장에서 재고를 찾았습니다.")
                print("재고 있는 매장:", stock_info)

            except Exception as e:
                print(f"재고 조회 중 오류 발생: {e}")
                continue

        # 매장 재고 조회 성공 후 처리
        if total_stores_count > 0:
            if max_stores_to_save > 0 and len(stock_info) > max_stores_to_save:
                stock_info_to_save = stock_info[:max_stores_to_save]
            else:
                stock_info_to_save = stock_info
            stock_info_str = '\n'.join(stock_info_to_save)
        else:
            stock_info_str = 'N/A'

        # 재고 현황 문자열 설정
        if total_stores_count == 0 and seoul_dosan_count == 0:
            stock_status = 'N/A'
        else:
            stock_status = f"{total_stores_count}({seoul_dosan_count})"
        
        # 공홈 등록 상태 업데이트
        if product_price == 'N/A' and stock_info_str == 'N/A':
            website_status = NOT_AVAILABLE
        elif product_price != 'N/A' and stock_info_str == 'N/A':
            website_status = ON_WEBSITE_NOT_SOLD  
        elif product_price == 'N/A' and stock_info_str != 'N/A':
            website_status = NOT_ON_WEBSITE_SOLD
        else:
            website_status = ON_WEBSITE_SOLD

        # 결과를 엑셀 파일에 쓰기
        sheet.cell(row=row_num, column=2, value=product_name)
        sheet.cell(row=row_num, column=3, value=product_price)
        sheet.cell(row=row_num, column=4, value=material)
        sheet.cell(row=row_num, column=5, value=size)
        sheet.cell(row=row_num, column=6, value=features)
        sheet.cell(row=row_num, column=7, value=description)
        sheet.cell(row=row_num, column=8, value=stock_status)
        sheet.cell(row=row_num, column=9, value=website_status)  # 열 추가 및 공홈 등록 상태 기록
        sheet.cell(row=row_num, column=10, value=stock_info_str)  # 재고 매장 정보 입력

        # 저장
        workbook.save(filename=target_path)
        print(f"제품 정보를 엑셀 파일에 저장했습니다.")

# 작업 완료 메시지 출력
print("작업이 완료되었습니다.")
