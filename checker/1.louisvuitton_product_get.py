from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import os

def create_driver():
    # 크롬 옵션 설정
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--disable-infobars')  # 정보 바 숨기기
    options.add_argument('--start-maximized')  # 창 최대화
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    options.add_argument('--disable-web-security')  # 웹 보안 비활성화
    options.add_argument('--allow-running-insecure-content')  # 안전하지 않은 콘텐츠 실행
    options.add_argument('--log-level=3')  # 로그 레벨 설정

    # 크롬드라이버 설정
    webdriver_service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=webdriver_service, options=options)
    return driver

# 시작 URL 및 시트 이름 매핑
start_urls = {
    "https://kr.louisvuitton.com/kor-kr/women/handbags/_/N-tfr7qdp": "여성 > 핸드백 전체",
    "https://kr.louisvuitton.com/kor-kr/women/travel/travel-bags/_/N-tdxhh3p": "여성 > 트래블 > 러기지 전체",
    "https://kr.louisvuitton.com/kor-kr/women/travel/travel-accessories/_/N-tbc3s3k": "여성 > 트래블 > 트래블 액세사리",
    "https://kr.louisvuitton.com/kor-kr/women/small-leather-goods/long-wallets/_/N-t1wnwc4o": "여성 > 지갑가죽소품 > 장지갑",
    "https://kr.louisvuitton.com/kor-kr/women/small-leather-goods/compact-wallets/_/N-ttz8v5o": "여성 > 지갑가죽소품 > 반 지갑",
    "https://kr.louisvuitton.com/kor-kr/women/small-leather-goods/chain-and-strap-wallets/_/N-t11y8dds": "여성 > 지갑가죽소품 > 체인지갑",
    "https://kr.louisvuitton.com/kor-kr/women/small-leather-goods/pouches/_/N-twzkl0p": "여성 > 지갑가죽소품 > 파우치",
    "https://kr.louisvuitton.com/kor-kr/women/small-leather-goods/card-holders-and-key-holders/_/N-tngffj4": "여성 > 지갑가죽소품 > 카드 홀더",
    "https://kr.louisvuitton.com/kor-kr/women/travel/travel-accessories/_/N-tbc3s3k": "여성 > 트래블 > 트래블 액세사리",
    "https://kr.louisvuitton.com/kor-kr/bags/for-men/all-bags/_/N-t1z0ff7q": "남자 > 핸드백 전체"
}

# 중복된 품번을 저장하기 위한 set 생성
unique_product_numbers_sheets = {name: set() for name in start_urls.values()}

def scroll_slowly(driver, scroll_pause_time=0.1, scroll_height=100):
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script(f"window.scrollBy(0, {scroll_height});")
        time.sleep(scroll_pause_time)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            print("Reached the end of the page during scrolling.")
            break
        last_height = new_height

def click_more_button(driver):
    try:
        more_button = driver.find_element(By.XPATH, "//span[contains(text(), '더보기')]")
        more_button.click()
        print("Clicked '더보기' button")
        time.sleep(2)
        return True
    except Exception as e:
        print(f"No more '더보기' button found: {e}")
        return False

def close_popup(driver):
    try:
        close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".lv-notifications__close.lv-button.-only-icon"))
        )
        close_button.click()
        print("Popup closed")
    except Exception as e:
        print("No popup found to close")

def process_url(driver, start_url, sheet_name):
    print(f"Processing {sheet_name} from URL: {start_url}")
    driver.get(start_url)
    time.sleep(5)  # 페이지 로딩 대기

    # Close popup if it appears
    close_popup(driver)

    last_product_ids = set()
    start_time = time.time()

    while True:
        scroll_slowly(driver)
        click_more_button(driver)

        # 현재 파싱한 제품 ID 확인
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        product_ids = set([h2_tag['id'].replace('product-', '') for h2_tag in soup.find_all('h2', {'id': re.compile(r'^product-')})])

        if product_ids == last_product_ids:
            elapsed_time = time.time() - start_time
            if elapsed_time >= 10:  # 10초 동안 새로운 품번이 없으면 다음 URL로 이동
                print(f"No new product IDs found for 10 seconds. Moving to the next URL.")
                break
        else:
            start_time = time.time()  # 새로운 품번이 추가되면 시작 시간 초기화
            last_product_ids = product_ids

    # 스크롤이 끝까지 내려간 후 페이지 소스 분석
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # h2 태그에서 제품 ID 추출
    product_ids = [h2_tag['id'].replace('product-', '') for h2_tag in soup.find_all('h2', {'id': re.compile(r'^product-')})]
    print(f"Found {len(product_ids)} product IDs")

    # 중복 제거를 위해 set에 품번 추가
    for product_id in product_ids:
        unique_product_numbers_sheets[sheet_name].add(product_id)

    # 데이터프레임 생성
    print(f"Creating DataFrame for {sheet_name}")
    df_product_numbers = pd.DataFrame(list(unique_product_numbers_sheets[sheet_name]), columns=["품번"])

    # 엑셀 파일에 쓰기
    output_excel_path = "루이비통.xlsx"
    if os.path.exists(output_excel_path):
        print(f"Appending data to existing Excel file for {sheet_name}")
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_product_numbers.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False, header=False)
            worksheet = writer.sheets[sheet_name]
            if worksheet.cell(row=1, column=1).value != '품번':  # 시트 첫 행에 '품번'이 없는지 확인
                worksheet.cell(row=1, column=1, value='품번')
    else:
        print(f"Creating new Excel file and writing data for {sheet_name}")
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w') as writer:
            df_product_numbers.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False, header=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.cell(row=1, column=1, value='품번')

def main():
    driver = create_driver()
    try:
        # 각 시작 URL에 대해 페이지를 탐색
        for start_url, sheet_name in start_urls.items():
            process_url(driver, start_url, sheet_name)
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    while True:
        main()
        print("Processing complete. Restarting in 5 seconds...")
        time.sleep(5)
