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
import customtkinter as ctk
import threading
from tkinter import Toplevel

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

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


class ProgressWindow(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.title("진행 상황")
        self.geometry("300x100")

        self.configure(bg='#2B2B2B')  # 다크모드 배경색 설정

        self.progress_bar = ctk.CTkProgressBar(self, width=250)
        self.progress_bar.pack(padx=20, pady=20)
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(self, text="데이터 수집 중...", fg_color='#2B2B2B', text_color='white')
        self.status_label.pack(padx=20, pady=10)

    def update_progress(self, value, status_text):
        self.progress_bar.set(value)
        self.status_label.configure(text=status_text)
        self.update_idletasks()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("KREAM 데이터 스크래퍼")
        self.geometry("600x700")  # 초기 윈도우 크기

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(8, weight=1)

        self.create_widgets()

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

        # 로그인 버튼
        self.login_button = ctk.CTkButton(self, text="로그인", command=self.login)
        self.login_button.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

        # 확인 버튼
        self.confirm_button = ctk.CTkButton(self, text="확인", command=self.confirm_login)
        self.confirm_button.grid(row=5, column=0, padx=20, pady=10, sticky="ew")
        self.confirm_button.grid_remove()

        # 데이터 수집 버튼
        self.scrape_button = ctk.CTkButton(self, text="데이터 수집", command=self.start_scraping)
        self.scrape_button.grid(row=6, column=0, padx=20, pady=10, sticky="ew")
        self.scrape_button.grid_remove()

        # 상태 메시지
        self.message_label = ctk.CTkLabel(self, text="")
        self.message_label.grid(row=7, column=0, padx=20, pady=10, sticky="ew")

        # 로그 텍스트 박스
        self.log_text = ctk.CTkTextbox(self, state="disabled")
        self.log_text.grid(row=8, column=0, padx=20, pady=10, sticky="nsew")

        self.progress_window = None

        # 창 크기 조정 이벤트 바인딩
        self.bind("<Configure>", self.on_resize)

    def on_resize(self, event):
        # 창 크기가 변경될 때 로그 텍스트 박스의 크기를 조정
        width = self.winfo_width() - 40  # 좌우 패딩 고려
        height = self.winfo_height() - 400  # 상단 위젯들의 높이를 고려하여 조정
        self.log_text.configure(width=width, height=height)

    def log(self, message):
        print(message)  # 콘솔에 출력
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update_idletasks()

    def login(self):
        self.driver = setup_driver(headless=False)
        self.driver.get("https://kream.co.kr/")

        try:
            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, LOGIN_BUTTON_CSS))
            )
            login_button.click()
            self.log("로그인 페이지로 이동했습니다.")
        except Exception as e:
            self.log(f"로그인 버튼 클릭 중 오류 발생: {e}")

        self.login_button.grid_remove()
        self.confirm_button.grid()
        self.message_label.configure(text="로그인 완료 후, 확인 버튼을 눌러주세요.")

    def confirm_login(self):
        self.confirm_button.grid_remove()
        self.scrape_button.grid()
        self.message_label.configure(text="로그인 확인 완료. 데이터 수집을 시작하세요.")

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
        self.progress_window = ProgressWindow(self)
        self.log("데이터 수집을 시작합니다...")

        threading.Thread(target=self.scrape_data, args=(brand, period, min_trades)).start()

    def scrape_data(self, brand, period, min_trades):
        # 현재 세션의 쿠키와 세션 데이터 가져오기
        cookies = self.driver.get_cookies()
        self.driver.quit()

        # 새로운 브라우저 설정
        self.driver = setup_driver(headless=False)

        # 쿠키와 세션 데이터 설정
        self.driver.get("https://kream.co.kr/")
        for cookie in cookies:
            self.driver.add_cookie(cookie)
        self.driver.refresh()

        detailed_products = scrape_data(self.driver, brand, self.update_progress, self.log, period, min_trades)
        self.progress_window.update_progress(1, "데이터 수집 완료")
        self.log(f"데이터 수집 완료. {brand}_detailed_products.xlsx 파일이 저장되었습니다.")
        self.scrape_button.configure(state="normal")
        self.progress_window.after(3000, self.progress_window.destroy)  # 3초 후 진행 창 닫기

    def update_progress(self, value):
        if self.progress_window:
            self.progress_window.update_progress(value, f"데이터 수집 중... {value:.0%}")


def setup_driver(headless=False):
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--disable-infobars')
    if headless:
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
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
        product_name = f"{title} / {subtitle}"
    except NoSuchElementException:
        product_name = "N/A"
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

    product_info = {'제품명': product_name, '모델번호': model_num}

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
                EC.element_to_be_clickable((By.XPATH, BUY_BUTTON_XPATH)))
            buy_button.click()
            time.sleep(5)
        except (TimeoutException, NoSuchElementException) as e:
            log_callback("즉시구매 버튼을 클릭하지 못했습니다. 대체 버튼을 시도합니다.")
            try:
                alt_buy_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ALT_BUY_BUTTON_CSS)))
                alt_buy_button.click()
                time.sleep(5)
            except (TimeoutException, NoSuchElementException) as e2:
                log_callback(f"대체 즉시구매 버튼을 클릭하지 못했습니다: {e2}")
                raise e2

        option_info = []
        option_elements = driver.find_elements(By.CSS_SELECTOR, OPTION_ELEMENTS_CSS)
        if not option_elements:
            try:
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
                log_callback("ONE SIZE 옵션을 찾을 수 없습니다.")
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
        # 수집한 제품 정보를 즉시 엑셀에 저장
        save_to_excel(product_info, f"{brand}_detailed_products.xlsx", log_callback, mode='a')

    except (TimeoutException, NoSuchElementException) as e:
        log_callback(f"구매 정보를 찾을 수 없습니다: {e}")
        return None

    return product_info


def scrape_data(driver, brand, progress_callback, log_callback, period, min_trades):
    initial_products = []
    detailed_products = []
    try:
        encoded_keyword = urllib.parse.quote(brand)
        url = f"https://kream.co.kr/search?keyword={encoded_keyword}&tab=products"
        driver.get(url)
        log_callback(f"'{brand}' 검색 결과 페이지로 이동했습니다.")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, SEARCH_RESULT_ITEM_CSS)))

        scroll_pause_time = 1
        scroll_height = 500
        last_height = driver.execute_script("return document.body.scrollHeight")

        start_time = time.time()  # 스크롤 시작 시간
        while True:
            driver.execute_script(f"window.scrollBy(0, {scroll_height});")
            time.sleep(scroll_pause_time)
            new_height = driver.execute_script("return document.body.scrollHeight")

            if new_height == last_height:
                if time.time() - start_time > 60:  # 1분 동안 새로운 제품이 로드되지 않으면 스크롤 중단
                    log_callback("1분 동안 새로운 제품이 로드되지 않았습니다. 스크롤을 중단합니다.")
                    break
            else:
                start_time = time.time()  # 새로운 제품이 로드되면 시간 초기화

            last_height = new_height

        log_callback("페이지 스크롤 완료")

        product_cards = driver.find_elements(By.CSS_SELECTOR, SEARCH_RESULT_ITEM_CSS)
        log_callback(f"총 {len(product_cards)}개의 제품을 찾았습니다.")

        for card in product_cards:
            product_info = get_initial_product_info(card)
            if product_info:
                initial_products.append(product_info)

        log_callback(f"빠른배송 가능한 제품 {len(initial_products)}개를 찾았습니다.")

        total_products = len(initial_products)
        for i, product in enumerate(initial_products):
            log_callback(f"제품 정보 수집 중: {product['품번']} ({i + 1}/{total_products})")
            detailed_info = get_detailed_product_info(driver, product['URL'], log_callback, brand, period, min_trades)
            if detailed_info:
                detailed_products.append(detailed_info)
                log_callback(f"상세 정보 수집 완료: {detailed_info['모델번호']}")
            else:
                log_callback(f"{product['품번']}에 대한 상세 정보를 수집하지 못했습니다. 다음 품번으로 넘어갑니다.")
            progress = (i + 1) / total_products
            progress_callback(progress)

        log_callback(f"총 {len(detailed_products)}개의 조건에 맞는 제품을 찾았습니다.")

    except Exception as e:
        log_callback(f"오류 발생: {str(e)}")

    return detailed_products


def save_to_excel(product_info, filename, log_callback, mode='w'):
    data = []
    product_name = product_info['제품명']
    model_num = product_info['모델번호']
    for option in product_info['옵션']:
        size = option['사이즈']
        price = option['가격']
        is_express = "빠른배송" if option['빠른배송'] else "일반배송"
        data.append([product_name, model_num, size, price, is_express])

    df = pd.DataFrame(data, columns=['제품명', '모델번호', '사이즈', '가격', '배송타입'])

    if mode == 'w':
        df.to_excel(filename, index=False)
    else:
        with pd.ExcelWriter(filename, mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, index=False, header=writer.sheets['Sheet1'].max_row == 0,
                        startrow=writer.sheets['Sheet1'].max_row)

    log_callback(f"{product_name} - 데이터가 {filename}에 추가 저장되었습니다.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
