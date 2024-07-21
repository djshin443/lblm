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
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import customtkinter as ctk
import threading
import json
import os

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# 변수 정의
LOGIN_URL = "https://kream.co.kr/login"
LOGOUT_BUTTON_CSS = 'a.top_link[href="/"]'  # 로그아웃 버튼의 CSS 셀렉터
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
BUY_BUTTON_CSS = 'button.btn_action .title'
OPTION_ELEMENTS_CSS = 'div.select_area ul.select_list li.select_item button.select_link.buy'
ONE_SIZE_CSS = 'button.select_link.buy'
SIZE_CSS = '.size'
PRICE_CSS = '.price'
EXPRESS_CSS = '.ico-express'

COOKIE_FILE = "cookies.json"
CACHE_FILE = "cache.json"
CACHE_VALIDITY_DAYS = 1  # 캐시 파일의 유효 기간을 1일로 설정


def save_cookies(driver, filepath):
    cookies = driver.get_cookies()
    for cookie in cookies:
        if 'expiry' not in cookie:
            cookie['expiry'] = int(time.time()) + 31536000  # 1년 후
    with open(filepath, 'w') as file:
        json.dump(cookies, file)


def load_cookies(driver, filepath):
    with open(filepath, 'r') as file:
        cookies = json.load(file)
        for cookie in cookies:
            driver.add_cookie(cookie)


def delete_cookies(filepath):
    if os.path.exists(filepath):
        os.remove(filepath)


def save_to_cache(data, filepath):
    if os.path.exists(filepath):
        existing_data = load_from_cache(filepath)
        data = existing_data + data
    with open(filepath, 'w') as file:
        json.dump(data, file)


def load_from_cache(filepath):
    if os.path.exists(filepath):
        with open(filepath, 'r') as file:
            return json.load(file)
    return []


def delete_cache(filepath):
    if os.path.exists(filepath):
        os.remove(filepath)


def is_cache_valid(filepath, validity_days):
    if os.path.exists(filepath):
        file_time = datetime.fromtimestamp(os.path.getmtime(filepath))
        if (datetime.now() - file_time).days < validity_days:
            return True
    return False


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("KREAM 데이터 스크래퍼")
        self.geometry("600x700")  # 초기 윈도우 크기

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(8, weight=1)

        self.create_widgets()

        # UI가 먼저 나오도록 하고, 비동기로 로그인 상태 확인
        self.after(100, self.check_initial_login_status)

    def create_widgets(self):
        # 헤더 라벨
        self.header_label = ctk.CTkLabel(self, text="KREAM 데이터 스크래퍼", font=("Arial", 20, "bold"))
        self.header_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="n")

        # 검색어 입력
        self.search_entry = ctk.CTkEntry(self, placeholder_text="검색어 입력")
        self.search_entry.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        # 스크래핑 기간 옵션
        self.period_var = ctk.StringVar(value="월")
        self.period_option = ctk.CTkOptionMenu(self, values=["일", "주", "월", "년"], variable=self.period_var)
        self.period_option.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        # 최소 거래 횟수 입력
        self.min_trades_entry = ctk.CTkEntry(self, placeholder_text="최소 거래 횟수 (기본값: 30)")
        self.min_trades_entry.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        # 확인 버튼
        self.confirm_button = ctk.CTkButton(self, text="확인", command=self.confirm_login)
        self.confirm_button.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        self.confirm_button.grid_remove()

        # 데이터 수집 버튼
        self.scrape_button = ctk.CTkButton(self, text="데이터 수집", command=self.start_scraping)
        self.scrape_button.grid(row=5, column=0, padx=20, pady=10, sticky="ew")
        self.scrape_button.grid_remove()

        # 캐시 초기화 버튼
        self.clear_cache_button = ctk.CTkButton(self, text="캐시 초기화", command=self.clear_cache)
        self.clear_cache_button.grid(row=6, column=0, padx=20, pady=10, sticky="ew")
        self.clear_cache_button.grid_remove()

        # 쿠키 삭제 버튼
        self.delete_cookies_button = ctk.CTkButton(self, text="로그인 정보 삭제", command=self.delete_cookies)
        self.delete_cookies_button.grid(row=7, column=0, padx=20, pady=10, sticky="ew")
        self.delete_cookies_button.grid_remove()

        # 진행 상황 라벨
        self.progress_label = ctk.CTkLabel(self, text="진행 상황: 0%", fg_color='#1E1E1E', corner_radius=8)
        self.progress_label.grid(row=8, column=0, padx=20, pady=(10, 5), sticky="ew")

        # 로그 텍스트 박스
        self.log_text = ctk.CTkTextbox(self, state="disabled", fg_color='#2B2B2B', corner_radius=8)
        self.log_text.grid(row=9, column=0, padx=20, pady=(5, 20), sticky="nsew")

        # 로그 텍스트 박스의 높이를 조정
        self.grid_rowconfigure(9, weight=1)  # 로그 텍스트 박스가 확장되도록 설정

        self.progress_window = None

        # 창 크기 조정 이벤트 바인딩
        self.bind("<Configure>", self.on_resize)

    def on_resize(self, event):
        width = self.winfo_width() - 40  # 좌우 패딩 고려
        height = max(100, self.winfo_height() - 400)  # 최소 높이 100px, 상단 위젯들의 높이를 고려하여 조정
        self.log_text.configure(width=width, height=height)

    def log(self, message):
        print(message)  # 콘솔에 출력
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.tag_add("highlight", "end-2l", "end-1l")
        self.log_text.tag_config("highlight", foreground="orange")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update_idletasks()

    def confirm_login(self):
        self.confirm_button.grid_remove()
        self.scrape_button.grid()
        self.clear_cache_button.grid()
        self.delete_cookies_button.grid()

        # 쿠키 저장
        self.after_login_save_cookies()

    def after_login_save_cookies(self):
        self.log("쿠키 정보를 저장합니다.")
        save_cookies(self.driver, COOKIE_FILE)

    def delete_cookies(self):
        delete_cookies(COOKIE_FILE)
        self.log("저장된 쿠키 정보가 삭제되었습니다.")
        self.scrape_button.grid_remove()
        self.clear_cache_button.grid_remove()
        self.delete_cookies_button.grid_remove()
        self.confirm_button.grid()

    def clear_cache(self):
        delete_cache(CACHE_FILE)
        self.log("캐시 파일이 초기화되었습니다.")

    def start_scraping(self):
        brand = self.search_entry.get()
        if not brand:
            self.log("검색어를 입력해주세요.")
            return

        period = self.period_var.get()
        min_trades = self.min_trades_entry.get()
        if not min_trades.isdigit():
            min_trades = 30
        else:
            min_trades = int(min_trades)

        self.scrape_button.configure(state="disabled")
        self.log("데이터 수집을 시작합니다...")

        # 새로운 브라우저 설정을 제거하고, 이미 설정된 브라우저를 재사용하도록 변경
        if not hasattr(self, 'driver') or self.driver is None:
            self.driver = setup_driver()

        threading.Thread(target=self.scrape_data, args=(brand, period, min_trades)).start()

    def scrape_data(self, brand, period, min_trades):
        try:
            # 캐시 파일의 유효성을 확인
            if is_cache_valid(CACHE_FILE, CACHE_VALIDITY_DAYS):
                initial_products = load_from_cache(CACHE_FILE)
                self.log("캐시 파일을 로드했습니다.")
            else:
                initial_products = self.scrape_initial_data(brand)
                save_to_cache(initial_products, CACHE_FILE)
                self.log("새로운 데이터를 스크래핑하여 캐시 파일을 생성했습니다.")

            detailed_products = self.scrape_detailed_data(initial_products, brand, period, min_trades)
            self.update_progress(1)
            self.log(f"데이터 수집 완료. {brand}_detailed_products.xlsx 파일이 저장되었습니다.")
            delete_cache(CACHE_FILE)  # 엑셀 저장 완료 후 캐시 파일 초기화
            self.log("캐시 파일이 초기화되었습니다.")
        except Exception as e:
            self.log(f"스크래핑 중 오류 발생: {e}")
        finally:
            self.scrape_button.configure(state="normal")
            # driver.quit() 호출을 제거하여 브라우저를 닫지 않음

    def update_progress(self, value):
        progress_percent = int(value * 100)
        self.progress_label.configure(text=f"진행 상황: {progress_percent}%")
        self.progress_label.update_idletasks()

    def check_initial_login_status(self):
        if os.path.exists(COOKIE_FILE):
            self.driver = setup_driver()
            self.log("저장된 쿠키 정보를 사용하여 자동 로그인 시도 중...")
            self.driver.get("https://kream.co.kr/")
            try:
                load_cookies(self.driver, COOKIE_FILE)
                self.driver.refresh()
                self.after(3000, self.check_login_status)
            except Exception:
                self.log("쿠키 파일이 유효하지 않습니다. 쿠키 파일을 초기화하고 다시 로그인합니다.")
                delete_cookies(COOKIE_FILE)
                self.driver.get(LOGIN_URL)
                self.confirm_button.grid()
        else:
            self.log("저장된 쿠키 정보가 없습니다. 수동 로그인이 필요합니다.")
            self.driver = setup_driver()
            self.driver.get(LOGIN_URL)
            self.confirm_button.grid()

    def check_login_status(self):
        try:
            # 로그인된 상태를 확인하기 위해 로그아웃 버튼을 찾습니다.
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, LOGOUT_BUTTON_CSS))
            )
            self.log("자동 로그인에 성공했습니다.")
            self.scrape_button.grid()
            self.clear_cache_button.grid()
            self.delete_cookies_button.grid()
        except TimeoutException:
            self.log("자동 로그인이 실패했습니다. 수동 로그인이 필요합니다.")
            self.driver.get(LOG인_URL)
            self.confirm_button.grid()
        except Exception as e:
            self.log(f"로그인 상태 확인 중 오류 발생: {e}")
            self.driver.get(LOGIN_URL)
            self.confirm_button.grid()

    def scrape_initial_data(self, brand):
        initial_products = []
        try:
            encoded_keyword = urllib.parse.quote(brand)
            url = f"https://kream.co.kr/search?keyword={encoded_keyword}&tab=products"
            self.driver.get(url)
            self.log(f"'{brand}' 검색 결과 페이지로 이동했습니다.")

            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, SEARCH_RESULT_ITEM_CSS)))

            scroll_pause_time = 1
            scroll_height = 500
            last_height = self.driver.execute_script("return document.body.scrollHeight")

            start_time = time.time()  # 스크롤 시작 시간
            while True:
                self.driver.execute_script(f"window.scrollBy(0, {scroll_height});")
                time.sleep(scroll_pause_time)
                new_height = self.driver.execute_script("return document.body.scrollHeight")

                # 스크롤이 멈추지 않았는지 확인
                if new_height == last_height:
                    if time.time() - start_time > 500:  # 1분 동안 새로운 제품이 로드되지 않으면 스크롤 중단
                        self.log("1분 동안 새로운 제품이 로드되지 않았습니다. 스크롤을 중단합니다.")
                        break
                else:
                    start_time = time.time()  # 새로운 제품이 로드되면 시간 초기화
                    last_height = new_height

                product_cards = self.driver.find_elements(By.CSS_SELECTOR, SEARCH_RESULT_ITEM_CSS)
                existing_product_urls = [product['URL'] for product in initial_products]  # 이미 수집된 제품 URL 목록
                for card in product_cards:
                    product_info = get_initial_product_info(card)
                    if product_info and product_info['URL'] not in existing_product_urls:
                        initial_products.append(product_info)
                        # 각 제품 정보를 캐시에 저장
                        save_to_cache(initial_products, CACHE_FILE)

                self.log(f"현재까지 수집된 제품 수: {len(initial_products)}")
                self.update_progress(len(initial_products) / 1000)  # 예시로 최대 1000개의 제품 수집 목표

                # 추가로 로드할 제품이 없으면 중단
                if len(product_cards) == 0:
                    break

            self.log(f"빠른배송 가능한 제품 {len(initial_products)}개를 찾았습니다.")

        except Exception as e:
            self.log(f"오류 발생: {str(e)}")

        return initial_products

    def scrape_detailed_data(self, initial_products, brand, period, min_trades):
        detailed_products = []
        total_products = len(initial_products)
        for i, product in enumerate(initial_products):
            self.log(f"제품 정보 수집 중: {product['품번']} ({i + 1}/{total_products})")
            detailed_info = get_detailed_product_info(self.driver, product['URL'], self.log, brand, period, min_trades)
            if detailed_info:
                detailed_products.append(detailed_info)
                self.log(f"상세 정보 수집 완료: {detailed_info['모델번호']}")
                save_to_excel(detailed_info, f"{brand}_detailed_products.xlsx", self.log, mode='a')
            else:
                self.log(f"{product['품번']}에 대한 상세 정보를 수집하지 못했습니다. 다음 품번으로 넘어갑니다.")
            progress = (i + 1) / total_products
            self.update_progress(progress)

        self.log(f"총 {len(detailed_products)}개의 조건에 맞는 제품을 찾았습니다.")
        return detailed_products


def setup_driver():
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--disable-infobars')
    options.add_argument('--start-maximized')
    options.add_argument('--log-level=3')
    options.add_argument('--silent')
    options.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
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


def get_detailed_product_info(driver, url, log_callback, brand, period, min_trades):
    driver.get(url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, DETAIL_BOX_CSS)))

    try:
        title = driver.find_element(By.CSS_SELECTOR, TITLE_CSS).text.strip()
        subtitle = driver.find_element(By.CSS_SELECTOR, SUBTITLE_CSS).text.strip()
        product_name_eng = title
        product_name_kor = subtitle
    except NoSuchElementException:
        product_name_eng = "N/A"
        product_name_kor = "N/A"
        log_callback("제품명을 찾을 수 없습니다.")

    try:
        model_num_element = driver.find_element(By.XPATH, MODEL_NUM_XPATH)
        model_num = model_num_element.text.strip()
    except NoSuchElementException:
        log_callback("모델번호를 찾을 수 없습니다.")
        return None

    try:
        time.sleep(5)
        more_button = driver.find_element(By.CSS_SELECTOR, MORE_BUTTON_CSS)
        more_button.click()
        time.sleep(2)
    except NoSuchElementException:
        log_callback("체결 내역 더보기 버튼을 찾을 수 없습니다.")
        return None

    trade_history = driver.find_elements(By.CSS_SELECTOR, TRADE_HISTORY_CSS)
    one_month_ago = datetime.now() - timedelta(days=30)
    if period == '일':
        one_month_ago = datetime.now() - timedelta(days=1)
    elif period == '주':
        one_month_ago = datetime.now() - timedelta(weeks=1)
    elif period == '월':
        one_month_ago = datetime.now() - timedelta(days=30)
    elif period == '년':
        one_month_ago = datetime.now() - timedelta(days=365)

    recent_trades = []

    for trade in trade_history:
        date_text = trade.find_element(By.CSS_SELECTOR, '.list_txt.is_active span').text
        if '빠른배송' in date_text:
            continue
        try:
            trade_date = datetime.strptime(date_text, '%y/%m/%d')
            if trade_date > one_month_ago:
                recent_trades.append(trade)
        except ValueError as e:
            log_callback(f"날짜 파싱 오류: {e}")
            continue

    if len(recent_trades) < min_trades:
        log_callback(f"최근 {period} 내 거래 내역이 {min_trades}개 미만입니다. (현재: {len(recent_trades)}개)")
        return None

    product_info = {'제품명(영문)': product_name_eng, '제품명(한글)': product_name_kor, '모델번호': model_num}

    try:
        time.sleep(5)
        close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, CLOSE_BUTTON_CSS)))
        close_button.click()
        time.sleep(1)
    except TimeoutException:
        log_callback("거래 내역 창 닫기 버튼을 찾을 수 없습니다.")
        return None

    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)

    try:
        try:
            time.sleep(5)
            buy_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, BUY_BUTTON_CSS)))
            buy_button.click()
            time.sleep(5)
        except (TimeoutException, NoSuchElementException) as e:
            log_callback("즉시구매 버튼을 클릭하지 못했습니다. 대체 버튼을 시도합니다.")
            raise e

        option_info = []
        option_elements = driver.find_elements(By.CSS_SELECTOR, OPTION_ELEMENTS_CSS)
        if not option_elements:
            try:
                one_size_element = driver.find_element(By.CSS_SELECTOR, ONE_SIZE_CSS)
                size = one_size_element.find_element(By.CSS_SELECTOR, SIZE_CSS).text.strip()
                price = one_size_element.find_element(By.CSS_SELECTOR, PRICE_CSS).text.strip()
                is_express = bool(one_size_element.find_elements(By.CSS_SELECTOR, EXPRESS_CSS))
                option_info.append({
                    '옵션(사이즈)': size,
                    '가격': price,
                    '빠른배송': is_express
                })
            except NoSuchElementException:
                log_callback("ONE SIZE 옵션을 찾을 수 없습니다.")
        else:
            for option in option_elements:
                try:
                    size = option.find_element(By.CSS_SELECTOR, SIZE_CSS).text.strip()
                    price = option.find_element(By.CSS_SELECTOR, PRICE_CSS).text.strip()
                    is_express = bool(option.find_elements(By.CSS_SELECTOR, EXPRESS_CSS))
                    option_info.append({
                        '옵션(사이즈)': size,
                        '가격': price,
                        '빠른배송': is_express
                    })
                except NoSuchElementException:
                    continue

        product_info['옵션'] = option_info

    except (TimeoutException, NoSuchElementException) as e:
        log_callback(f"구매 정보를 찾을 수 없습니다: {e}")
        return None

    return product_info


def save_to_excel(product_info, filename, log_callback, mode='w'):
    data = []
    product_name_eng = product_info['제품명(영문)']
    product_name_kor = product_info['제품명(한글)']
    model_num = product_info['모델번호']
    for option in product_info['옵션']:
        size = option['옵션(사이즈)']
        price = option['가격']
        is_express = "빠른배송" if option['빠른배송'] else "일반배송"
        data.append([product_name_eng, product_name_kor, model_num, size, price, is_express])

    df = pd.DataFrame(data, columns=['제품명(영문)', '제품명(한글)', '모델번호', '옵션(사이즈)', '가격', '배송타입'])

    try:
        # 엑셀 파일이 이미 존재하는 경우, 기존 파일에 데이터를 추가
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # 존재하는 시트의 다음 행에 데이터 추가
            startrow = writer.sheets['Sheet1'].max_row
            df.to_excel(writer, index=False, header=False, startrow=startrow)
    except FileNotFoundError:
        # 엑셀 파일이 존재하지 않는 경우, 새로운 파일 생성
        df.to_excel(filename, index=False, header=True)

    log_callback(f"{product_name_eng} - 데이터가 {filename}에 추가 저장되었습니다.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
