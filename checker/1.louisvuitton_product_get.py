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
    "https://kr.louisvuitton.com/kor-kr/women/accessories/belts/_/N-ty346sh": "여성 > 액세서리 > 벨트" ,
    "https://kr.louisvuitton.com/kor-kr/women/accessories/silk-squares-and-bandeaus/_/N-tte63om": "여성 > 액세서리 > 실크 스퀘어 및 방도",
    "https://kr.louisvuitton.com/kor-kr/women/accessories/scarves-and-shawls/_/N-toyetr6": "여성 > 액세서리 > 숄 앤 스툴 ",
    "https://kr.louisvuitton.com/kor-kr/women/accessories/scarves/_/N-tmr7ugu": "여성 > 액세서리 > 스카프",
    "https://kr.louisvuitton.com/kor-kr/women/accessories/hair-accessories/_/N-t1ydxwto": "여성 > 액세서리 > 헤어 액세서리",
    "https://kr.louisvuitton.com/kor-kr/women/accessories/key-holders-and-bag-charms/_/N-tui7c0x": "여성 > 액세서리 > 키홀더 n 백참",
    
    "https://kr.louisvuitton.com/kor-kr/women/fashion-jewelry/bracelets/_/N-t7sxnbc": "여성 > 패션 주얼리 > 팔찌",
    "https://kr.louisvuitton.com/kor-kr/women/fashion-jewelry/rings/_/N-t175idqj": "여성 > 패션 주얼리 > 반지",
    "https://kr.louisvuitton.com/kor-kr/women/fashion-jewelry/earrings/_/N-t1fwp6ww": "여성 > 패션 주얼리 > 귀걸이",
    "https://kr.louisvuitton.com/kor-kr/women/fashion-jewelry/necklaces-and-pendants/_/N-tguaa6b": "여성 > 패션 주얼리 > 목걸이 n 펜던트",

    "https://kr.louisvuitton.com/kor-kr/women/shoes/sneakers/_/N-t13nbknz": "여성 > 슈즈 > 스니커즈",
    "https://kr.louisvuitton.com/kor-kr/women/shoes/boots-and-ankle-boots/_/N-t1is1ifz": "여성 > 슈즈 > 부츠 n 부티",
    "https://kr.louisvuitton.com/kor-kr/women/shoes/loafers-and-ballerinas/_/N-t18l8jj3": "여성 > 슈즈 > 로퍼 n 발레리나",
    "https://kr.louisvuitton.com/kor-kr/women/shoes/platform-shoes/_/N-twh7u2z": "여성 > 슈즈 > 플랫폼 슈즈",
    "hhttps://kr.louisvuitton.com/kor-kr/women/shoes/mules-and-slides/_/N-t1lkbycc": "여성 > 슈즈 > 뮬 n 슬라이드",
    "https://kr.louisvuitton.com/kor-kr/women/shoes/sandals-and-espadrilles/_/N-t1ey0zdr": "여성 > 슈즈 > 샌들",
    "https://kr.louisvuitton.com/kor-kr/women/shoes/pumps/_/N-t1wpeudq": "여성 > 슈즈 > 펌프스",

    "https://kr.louisvuitton.com/kor-kr/bags/for-men/all-bags/_/N-t1z0ff7q": "남성 > 핸드백 전체",

    "https://kr.louisvuitton.com/kor-kr/men/accessories/belts/_/N-t1g9dx5w": "남성 > 액세서리 > 벨트",
    "https://kr.louisvuitton.com/kor-kr/men/accessories/scarves/_/N-t186w8xz": "남성 > 액세서리 > 스카프",
    "https://kr.louisvuitton.com/kor-kr/men/accessories/hats-and-gloves/_/N-t1fq44ws": "남성 > 액세서리 > 모자 n 장갑",
    "https://kr.louisvuitton.com/kor-kr/men/accessories/ties-and-pocket-squares/_/N-tv1f2za": "남성 > 액세서리 > 타이 n 포켓 스퀘어",
    "https://kr.louisvuitton.com/kor-kr/men/accessories/key-holders-and-bag-charms/_/N-tzumbbt": "남성 > 액세서리 > 키홀더 n 백참",

    "https://kr.louisvuitton.com/kor-kr/men/fashion-jewelry/necklaces-and-pendants/_/N-t79dl9r": "남성 > 패션 주얼리 > 목걸이 n 펜던트",
    "https://kr.louisvuitton.com/kor-kr/men/fashion-jewelry/bracelets/_/N-t17pp1v0": "남성 > 패션 주얼리 > 팔찌",
    "https://kr.louisvuitton.com/kor-kr/men/fashion-jewelry/rings/_/N-t1l6jm7c": "남성 > 패션 주얼리 > 반지",
    "https://kr.louisvuitton.com/kor-kr/men/fashion-jewelry/earrings/_/N-t1nhj986": "남성 > 패션 주얼리 > 귀걸이",

    "https://kr.louisvuitton.com/kor-kr/men/shoes/sneakers/_/N-t1fud07g": "남성 > 슈즈 > 스니커즈",
    "https://kr.louisvuitton.com/kor-kr/men/shoes/loafers-and-moccasins/_/N-t1x4h6di": "남성 > 슈즈 > 로퍼 n 모카신",
    "https://kr.louisvuitton.com/kor-kr/men/shoes/lace-ups-and-buckles-shoes/_/N-tra4xg2": "남성 > 슈즈 > 레이스업",
    "https://kr.louisvuitton.com/kor-kr/men/shoes/boots/_/N-tp5lgyh": "남성 > 슈즈 > 부츠",
    "https://kr.louisvuitton.com/kor-kr/men/shoes/sandals/_/N-t1c9fsss": "남성 > 슈즈 > 샌들",

    "https://kr.louisvuitton.com/kor-kr/men/wallets-and-small-leather-goods/long-wallets/_/N-th7ik7d": "남성 > 지갑가죽소품 > 장지갑",
    "https://kr.louisvuitton.com/kor-kr/men/wallets-and-small-leather-goods/compact-wallets/_/N-teruo45": "남성 > 지갑가죽소품 > 반지갑",
    "https://kr.louisvuitton.com/kor-kr/men/wallets-and-small-leather-goods/cardholders-and-passport-cases/_/N-t8v1ks": "남성 > 지갑가죽소품 > 카드 지갑 n 여권 지갑",
    "https://kr.louisvuitton.com/kor-kr/men/wallets-and-small-leather-goods/wearable-wallets/_/N-t1yhvr0b": "남성 > 지갑가죽소품 > 웨어러블 가죽 소품",
    "https://kr.louisvuitton.com/kor-kr/men/wallets-and-small-leather-goods/pouches/_/N-t1202gxy": "남성 > 지갑가죽소품 > 파우치",

    "https://kr.louisvuitton.com/kor-kr/men/bags/leather-bags-selection/_/N-t1gza3m5": "남성 > 가방 > 레더백 셀렉션",
    "https://kr.louisvuitton.com/kor-kr/men/bags/iconic-monogram-bags/_/N-taeo5qx": "남성 > 가방 > 아이코닉 모노그램 백",
    "https://kr.louisvuitton.com/kor-kr/men/bags/damier-signature/_/N-tn65vv5": "남성 > 가방 > 다미에 시그니처",
    "https://kr.louisvuitton.com/kor-kr/men/bags/backpacks/_/N-t7ipiak": "남성 > 가방 > 백팩",
    "https://kr.louisvuitton.com/kor-kr/men/bags/crossbody-bags/_/N-t1bmxkvj": "남성 > 가방 > 크로스 백",
    "https://kr.louisvuitton.com/kor-kr/men/bags/business-bags/_/N-t7dnshy": "남성 > 가방 > 비즈니스 백",
    "https://kr.louisvuitton.com/kor-kr/men/bags/totes/_/N-t1ixjs8n": "남성 > 가방 > 토트 백",
    "https://kr.louisvuitton.com/kor-kr/men/bags/small-bags-and-belt-bags/_/N-t1pizyia": "남성 > 가방 > 스몰 백 n 벨트백",
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
            if elapsed_time >= 60:  # 10초 동안 새로운 품번이 없으면 다음 URL로 이동
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
        print("All URLs have been processed. The program will now exit.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
