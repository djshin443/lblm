import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import urllib.parse
from datetime import datetime, timedelta
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# 변수 정의
LOGIN_BUTTON_CSS = 'a.top_link[href="/login"]'
SEARCH_RESULT_ITEM_CSS = '.search_result_item'
QUICK_DELIVERY_TAG_CSS = '.tag.display_tag_item .tag_text'
PRODUCT_NAME_CSS = '.product_info_product_name .name'
PRODUCT_URL_CSS = 'a'
DETAIL_BOX_CSS = '.detail-box'
TITLE_CSS = '.main-title-container .title'
SUBTITLE_CSS = '.main-title-container .sub-title'
MODEL_NUM_XPATH = "//div[@class='detail-box']//div[contains(text(), '모델번호')]/following-sibling::div[@class='product_info']"
MORE_BUTTON_CSS = 'a[data-v-420a5cda][data-v-32864b0e].btn.outlinegrey.full.medium'
TRADE_HISTORY_CSS = '.body_list'
CLOSE_BUTTON_CSS = '#wrap > div.layout__main--without-search.container.detail.lg > div.content > div.column_bind > div:nth-child(2) > div > div.layer_market_price.layer.lg > div.layer_container > a > svg'
BUY_BUTTON_XPATH = '/html/body/div/div/div/div[3]/div[1]/div[1]/div[2]/div/div[1]/div[5]/div/button[1]/strong'
ALT_BUY_BUTTON_CSS = 'button.btn_action .title'
OPTION_ELEMENTS_CSS = 'div.select_area ul.select_list li.select_item button.select_link.buy'
ONE_SIZE_CSS = 'button.select_link.buy'
SIZE_CSS = '.size'
PRICE_CSS = '.price'
EXPRESS_CSS = '.ico-express'

def setup_driver():
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--disable-infobars')
    # options.add_argument('--headless')
    options.add_argument('--start-maximized')  # 창 최대화
    options.add_argument('--log-level=3')
    options.add_argument('--silent')  # Suppress all console output from Chrome
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def get_initial_product_info(card):
    product_info = {}
    try:
        tag_text_element = card.find_element(By.CSS_SELECTOR, QUICK_DELIVERY_TAG_CSS)
        if "빠른배송" not in tag_text_element.text:
            return None
    except NoSuchElementException:
        return None

    try:
        model_num = card.find_element(By.CSS_SELECTOR, PRODUCT_NAME_CSS).text.strip()
        product_info['품번'] = model_num
    except NoSuchElementException:
        return None

    try:
        url = card.find_element(By.CSS_SELECTOR, PRODUCT_URL_CSS).get_attribute('href')
        product_info['URL'] = url
    except NoSuchElementException:
        product_info['URL'] = "N/A"

    return product_info


def get_detailed_product_info(driver, url):
    driver.get(url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, DETAIL_BOX_CSS)))

    # 제품명 추출
    try:
        title = driver.find_element(By.CSS_SELECTOR, TITLE_CSS).text.strip()
        subtitle = driver.find_element(By.CSS_SELECTOR, SUBTITLE_CSS).text.strip()
        product_name = f"{title} / {subtitle}"
    except NoSuchElementException:
        product_name = "N/A"
        print("제품명을 찾을 수 없습니다.")

    # 모델번호 추출
    try:
        model_num_element = driver.find_element(By.XPATH, MODEL_NUM_XPATH)
        model_num = model_num_element.text.strip()
    except NoSuchElementException:
        print("모델번호를 찾을 수 없습니다.")
        return None

    # 거래 내역 확인
    try:
        time.sleep(5)  # 대기 시간 추가
        more_button = driver.find_element(By.CSS_SELECTOR, MORE_BUTTON_CSS)
        more_button.click()
        time.sleep(2)
    except NoSuchElementException:
        print("체결 내역 더보기 버튼을 찾을 수 없습니다.")
        return None

    trade_history = driver.find_elements(By.CSS_SELECTOR, TRADE_HISTORY_CSS)
    one_month_ago = datetime.now() - timedelta(days=30)
    recent_trades = []

    for trade in trade_history:
        date_text = trade.find_element(By.CSS_SELECTOR, '.list_txt.is_active span').text
        print(f"거래 날짜 텍스트: {date_text}")  # 디버그 메시지
        if '빠른배송' in date_text:
            continue
        try:
            trade_date = datetime.strptime(date_text, '%y/%m/%d')
            print(f"파싱된 거래 날짜: {trade_date}")  # 디버그 메시지
            if trade_date > one_month_ago:
                recent_trades.append(trade)
        except ValueError as e:
            print(f"날짜 파싱 오류: {e}")  # 디버그 메시지
            continue

    print(f"최근 거래 개수: {len(recent_trades)}")  # 디버그 메시지
    if len(recent_trades) < 30:
        print(f"최근 한 달 내 거래 내역이 30개 미만입니다. (현재: {len(recent_trades)}개)")
        return None

    product_info = {'제품명': product_name, '모델번호': model_num}

    try:
        time.sleep(5)  # 대기 시간 추가
        # 지정한 CSS 셀렉터를 사용하여 팝업 닫기
        close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, CLOSE_BUTTON_CSS)))
        print("거래 내역 팝업을 닫습니다.")  # 디버그 메시지
        close_button.click()
        time.sleep(1)
    except TimeoutException:
        print("거래 내역 창 닫기 버튼을 찾을 수 없습니다.")
        return None

    # 스크롤을 맨 위로 올리기
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)

    try:
        # 즉시구매 버튼을 확실히 클릭할 수 있도록 여러 번 시도
        try:
            time.sleep(5)  # 대기 시간 추가
            buy_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, BUY_BUTTON_XPATH)))
            print("즉시구매 버튼을 클릭합니다.")  # 디버그 메시지
            buy_button.click()
            time.sleep(5)  # 대기 시간 추가
        except (TimeoutException, NoSuchElementException) as e:
            print("즉시구매 버튼을 클릭하지 못했습니다. 대체 버튼을 시도합니다.")
            try:
                alt_buy_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ALT_BUY_BUTTON_CSS)))
                print("대체 즉시구매 버튼을 클릭합니다.")  # 디버그 메시지
                alt_buy_button.click()
                time.sleep(5)  # 대기 시간 추가
            except (TimeoutException, NoSuchElementException) as e2:
                print(f"대체 즉시구매 버튼을 클릭하지 못했습니다: {e2}")
                raise e2

        # 옵션 정보 추출
        option_info = []
        option_elements = driver.find_elements(By.CSS_SELECTOR, OPTION_ELEMENTS_CSS)
        if not option_elements:
            try:
                # ONE SIZE 처리
                one_size_element = driver.find_element(By.CSS_SELECTOR, ONE_SIZE_CSS)
                size = one_size_element.find_element(By.CSS_SELECTOR, SIZE_CSS).text.strip()
                price = one_size_element.find_element(By.CSS_SELECTOR, PRICE_CSS).text.strip()
                is_express = bool(one_size_element.find_elements(By.CSS_SELECTOR, EXPRESS_CSS))
                option_info.append({
                    '사이즈': size,
                    '가격': price,
                    '빠른배송': is_express
                })
            except NoSuchElementException:
                print("ONE SIZE 옵션을 찾을 수 없습니다.")
        else:
            for option in option_elements:
                try:
                    size = option.find_element(By.CSS_SELECTOR, SIZE_CSS).text.strip()
                    price = option.find_element(By.CSS_SELECTOR, PRICE_CSS).text.strip()
                    is_express = bool(option.find_elements(By.CSS_SELECTOR, EXPRESS_CSS))
                    option_info.append({
                        '사이즈': size,
                        '가격': price,
                        '빠른배송': is_express
                    })
                except NoSuchElementException:
                    continue

        product_info['옵션'] = option_info

    except (TimeoutException, NoSuchElementException) as e:
        print(f"구매 정보를 찾을 수 없습니다: {e}")
        return None

    return product_info


def scrape_data(brand):
    driver = setup_driver()
    initial_products = []
    detailed_products = []
    try:
        encoded_keyword = urllib.parse.quote(brand)
        url = f"https://kream.co.kr/search?keyword={encoded_keyword}&tab=products"
        driver.get(url)
        print(f"'{brand}' 검색 결과 페이지로 이동했습니다.")

        try:
            login_button = driver.find_element(By.CSS_SELECTOR, LOGIN_BUTTON_CSS)
            login_button.click()
            print("로그인 버튼을 클릭했습니다.")
        except Exception as e:
            print(f"로그인 버튼 클릭 중 오류 발생: {e}")

        input("로그인을 완료한 후 엔터 키를 눌러 주세요...")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, SEARCH_RESULT_ITEM_CSS)))

        # 동적 컨텐츠를 위해 스크롤을 천천히 내리는 코드 추가
        scroll_pause_time = 1
        scroll_height = 500
        last_height = driver.execute_script("return document.body.scrollHeight")

        while True:
            driver.execute_script(f"window.scrollBy(0, {scroll_height});")
            time.sleep(scroll_pause_time)
            new_height = driver.execute_script("return document.body.scrollHeight")

            if new_height == last_height:
                break

            last_height = new_height

        product_cards = driver.find_elements(By.CSS_SELECTOR, SEARCH_RESULT_ITEM_CSS)
        print(f"총 {len(product_cards)}개의 제품을 찾았습니다.")

        for card in product_cards:
            product_info = get_initial_product_info(card)
            if product_info:
                initial_products.append(product_info)

        print(f"빠른배송 가능한 제품 {len(initial_products)}개를 찾았습니다.")

        for product in initial_products:
            detailed_info = get_detailed_product_info(driver, product['URL'])
            if detailed_info:
                detailed_products.append(detailed_info)
                print(f"상세 정보 수집 완료: {detailed_info['모델번호']}")
            else:
                print(f"{product['품번']}에 대한 상세 정보를 수집하지 못했습니다. 다음 품번으로 넘어갑니다.")

        print(f"총 {len(detailed_products)}개의 조건에 맞는 제품을 찾았습니다.")

    except Exception as e:
        print(f"오류 발생: {str(e)}")
    finally:
        driver.quit()

    return detailed_products


def save_to_excel(products, filename):
    if products:
        data = []
        for product in products:
            product_name = product['제품명']
            model_num = product['모델번호']
            for option in product['옵션']:
                size = option['사이즈']
                price = option['가격']
                is_express = "빠른배송" if option['빠른배송'] else "일반배송"
                data.append([product_name, model_num, size, price, is_express])

        df = pd.DataFrame(data, columns=['제품명', '모델번호', '사이즈', '가격', '배송타입'])
        df.to_excel(filename, index=False)
        print(f"데이터가 {filename}에 저장되었습니다.")
    else:
        print("수집된 제품 정보가 없습니다.")


if __name__ == "__main__":
    brand = input("브랜드 이름을 입력하세요: ")
    detailed_products = scrape_data(brand)
    save_to_excel(detailed_products, f"{brand}_detailed_products.xlsx")
