from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import re
import time

# 크롬 옵션 설정
options = Options()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument('--disable-infobars')  # 정보 바 숨기기
options.add_argument('--headless')  # 브라우저를 표시하지 않음
options.add_argument('--window-size=1920,1080')  # 윈도우 크기 설정
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
options.add_argument('--disable-web-security')  # 웹 보안 비활성화
options.add_argument('--allow-running-insecure-content')  # 안전하지 않은 콘텐츠 실행 
options.add_argument('--log-level=3')  # 로그 레벨 설정

# 크롬드라이버 자동 업데이트 및 설정
webdriver_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=webdriver_service, options=options)

# 시작 URL 및 시트 이름 매핑
start_urls = {
    "https://kr.louisvuitton.com/kor-kr/women/handbags/_/N-tfr7qdp": "여자 > 핸드백",
    "https://kr.louisvuitton.com/kor-kr/women/shoes/all-shoes/_/N-t1mcbujj": "여자 > 슈즈",
}

# 중복된 품번을 저장하기 위한 set 생성
unique_product_numbers_sheets = {name: set() for name in start_urls.values()}

# 각 시작 URL에 대해 페이지를 탐색
for start_url, sheet_name in start_urls.items():
    page = 1
    while True:
        try:
            if page == 1:
                url = start_url
            else:
                url = f"{start_url}?page={page}"
            
            driver.get(url)
            time.sleep(5)  # 동적 컨텐츠 로딩 대기
            
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            
            # URL에서 품번 추출
            product_urls = re.findall(r'"url":"https://kr\.louisvuitton\.com/kor-kr/products/([^"]+)"', str(soup))
            
            if product_urls:
                # 중복 제거를 위해 set에 품번 추가
                for product_url in product_urls:
                    product_number = product_url.split('/')[-1]  # 품번
                    unique_product_numbers_sheets[sheet_name].add(product_number)
                
                print(f"Processing page {page} for URL: {start_url}")
                page += 1  # 다음 페이지로
            else:
                print(f"No more product numbers or URLs found for {start_url}, or reached the last page.")
                break
        except Exception as e:
            print(f"An error occurred for {start_url}: {e}")
            break

driver.quit()  # WebDriver 종료

# 엑셀 파일에 품번을 저장
output_excel_path = "product_info.xlsx"

with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    for sheet_name, product_numbers in unique_product_numbers_sheets.items():
        # 데이터프레임 생성
        df_product_numbers = pd.DataFrame(list(product_numbers), columns=["품번"])
        
        # 첫 번째 행(인덱스 행)에 '품번'을 추가하지 않고 데이터프레임 기록
        df_product_numbers.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False, header=False)
        
        # 작업 중인 워크시트 객체를 가져옴
        worksheet = writer.sheets[sheet_name]
        
        # 첫 번째 행 첫 번째 열(A1)에 '품번'을 한 번만 추가
        worksheet.write('A1', '품번')
