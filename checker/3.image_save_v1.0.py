import os
import re
import requests
import time
import logging
from random import uniform
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException

# Setup logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# File handler for error logging
log_file_path = 'combined_log.txt'
file_handler = logging.FileHandler(log_file_path)
file_handler.setLevel(logging.INFO)
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

# Additional logger for image download info
download_logger = logging.getLogger('download_logger')
download_logger.setLevel(logging.INFO)
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
download_logger.addHandler(console_handler)
download_logger.addHandler(file_handler)  # Add file handler to download logger as well

# Suppress unnecessary stack trace log
for logger_name in ['selenium.webdriver.remote.remote_connection', 'urllib3.connectionpool']:
    log = logging.getLogger(logger_name)
    log.setLevel(logging.WARNING)

def setup_driver():
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--disable-infobars')
    # options.add_argument('--headless')
    options.add_argument('--start-maximized')  # 창 최대화
    options.add_argument('--log-level=3')
    options.add_argument('--silent')  # Suppress all console output from Chrome
    options.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def download_image(image_url, folder, filename, max_retries=3, timeout=60):
    if not os.path.exists(folder):
        os.makedirs(folder)

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
    }

    retries = 0
    while retries < max_retries:
        try:
            response = requests.get(image_url, headers=headers, timeout=timeout)
            response.raise_for_status()

            file_path = os.path.join(folder, f'{filename}.png')
            file_counter = 1
            base_file_path = file_path

            while os.path.exists(file_path):
                file_path = f"{base_file_path}_{file_counter}.png"
                file_counter += 1

            with open(file_path, 'wb') as file:
                file.write(response.content)

            download_logger.info(f"Downloaded: {file_path}")
            with open(log_file_path, 'a') as log_file:
                log_file.write(f"Success: {file_path}\n")
            return True
        except requests.exceptions.RequestException as e:
            download_logger.error(f"Failed to download {image_url}")
            download_logger.error(f"Error: {e}")
            retries += 1
            download_logger.info(f"Retrying download... (Attempt {retries}/{max_retries})")
            time.sleep(5)  # Wait for 5 seconds before retrying

    download_logger.error(f"Failed to download {image_url} after {max_retries} retries")
    with open(log_file_path, 'a') as log_file:
        log_file.write(f"Failed: {image_url}\n")
    return False

def get_largest_image_url(srcset):
    url_pattern = re.compile(r'(https?:\/\/[^\s,]+)\s+(\d+)w')
    urls_with_sizes = url_pattern.findall(srcset)
    if not urls_with_sizes:
        return None
    largest_url = max(urls_with_sizes, key=lambda x: int(x[1]))
    return largest_url[0]  # Return only the URL part

def scroll_down_slowly(driver, scroll_pause_time=1, scroll_increment=500):
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        for i in range(0, last_height, scroll_increment):
            driver.execute_script(f"window.scrollBy(0, {scroll_increment});")
            time.sleep(scroll_pause_time)

        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def search_product(driver, product_code):
    try:
        search_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".lv-header__utility-label.list-label-s"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
        time.sleep(uniform(5, 10))  # Allow time for scrolling and random wait
        search_button.click()
        time.sleep(uniform(5, 10))  # Random wait after click

        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#searchHeaderInput"))
        )
        search_input.clear()
        time.sleep(uniform(5, 10))  # Random wait before typing
        search_input.send_keys(product_code)
        search_input.send_keys(Keys.ENTER)
        time.sleep(uniform(5, 10))  # Random wait after typing

        # 팝업 닫기
        close_popup(driver)

        # 검색 결과 확인
        no_results = driver.find_elements(By.CSS_SELECTOR, ".lv-no-search-results")
        if no_results:
            logger.warning(f"No search results found for product code: {product_code}")
            with open(log_file_path, 'a') as log_file:
                log_file.write(f"No product found: {product_code}\n")
            return False
    except ElementClickInterceptedException as e:
        logger.error(f"검색 중 오류 발생: {e}")
        with open(log_file_path, 'a') as log_file:
            log_file.write(f"Search error (ElementClickIntercepted): {product_code}\n")
        return False
    except TimeoutException as e:
        logger.error(f"Timeout during search: {e}")
        with open(log_file_path, 'a') as log_file:
            log_file.write(f"Search error (Timeout): {product_code}\n")
        return False
    except Exception as e:
        logger.error(f"검색 중 오류 발생: {e}")
        with open(log_file_path, 'a') as log_file:
            log_file.write(f"Search error (General): {product_code}\n")
        return False
    return True

def close_popup(driver):
    try:
        close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".lv-notifications__close.lv-button.-only-icon"))
        )
        close_button.click()
        logger.info("Popup closed")
    except TimeoutException:
        logger.info("No popup found to close")

def go_to_homepage(driver):
    driver.get("https://kr.louisvuitton.com/")
    time.sleep(uniform(5, 10))  # Wait for the homepage to load

def get_image_urls(driver, product_code):
    logger.info(f"Searching images for product code: {product_code}")
    if not search_product(driver, product_code):
        logger.info(f"No product found for product code: {product_code}")
        go_to_homepage(driver)
        return []

    scroll_down_slowly(driver, scroll_pause_time=2)
    time.sleep(5)

    image_elements = driver.find_elements(By.CSS_SELECTOR, "img[srcset], img[data-srcset]")
    logger.info(f"Found {len(image_elements)} image elements")
    image_urls = []
    seen_urls = set()
    first_product_code = None

    for idx, image_element in enumerate(image_elements):
        logger.info(f"Processing image element {idx + 1}")
        srcset = image_element.get_attribute("srcset") or image_element.get_attribute("data-srcset")
        fetch_priority = image_element.get_attribute("fetchpriority")

        if srcset:
            logger.info(f"Found srcset: {srcset}")
            largest_image_url = get_largest_image_url(srcset)
            if largest_image_url and largest_image_url not in seen_urls:
                match = re.search(r'--([A-Z0-9]{6,})', largest_image_url)
                if match:
                    image_product_code = match.group(1)
                    if first_product_code is None:
                        first_product_code = image_product_code
                    if image_product_code == first_product_code:
                        if fetch_priority == "high":
                            logger.info("Image has high fetch priority, waiting before processing...")
                            time.sleep(10)  # Wait longer for high priority images to load
                        image_urls.append(largest_image_url)
                        seen_urls.add(largest_image_url)
                        logger.info(f"Selected image URL: {largest_image_url}")
                else:
                    logger.info(f"No matching product code found in image URL: {largest_image_url}")
            else:
                logger.info(f"Skipping duplicate or no valid image URL found")
        else:
            logger.info("No srcset attribute found")

    if not image_urls:
        logger.warning(f"No images found for product code: {product_code}")
    else:
        logger.info(f"Found {len(image_urls)} image URLs for product code: {product_code}")

    go_to_homepage(driver)
    return image_urls

def read_product_codes_from_excel(file_path):
    logger.info(f"Reading product codes from Excel file: {file_path}")
    workbook = load_workbook(filename=file_path)
    product_codes = {}

    for sheet in workbook:
        codes = [cell.value for cell in sheet['A'][1:] if cell.value]
        product_codes[sheet.title] = codes

    return product_codes

def sanitize_folder_name(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

excel_file_path = "./루이비통.xlsx"
all_product_codes = read_product_codes_from_excel(excel_file_path)

def main():
    while True:
        try:
            driver = setup_driver()
            go_to_homepage(driver)

            for sheet_name, product_codes in all_product_codes.items():
                sanitized_sheet_name = sanitize_folder_name(sheet_name)
                for product_code in product_codes:
                    sanitized_product_code = sanitize_folder_name(str(product_code))
                    try:
                        image_urls = get_image_urls(driver, product_code)
                        if not image_urls:
                            continue

                        folder_path = os.path.join("image", sanitized_sheet_name, sanitized_product_code)
                        download_logger.info(f"Processing product code: {product_code}")
                        
                        for idx, url in enumerate(image_urls):
                            logger.info(f"Downloading image: {url}")
                            success = download_image(url, folder_path, f"{sanitized_product_code}_image_{idx + 1}")
                            if success:
                                logger.error(f"Download successful for {url}")
                            else:
                                logger.error(f"Download failed for {url}")
                            wait_time = uniform(30, 60)
                            logger.info(f"Waiting for {wait_time:.2f} seconds before next download...")
                            time.sleep(wait_time)  # Wait for 30 to 60 seconds between each download
                        go_to_homepage(driver)
                    except Exception as e:
                        logger.error(f"Error occurred while processing product code {product_code}: {e}")
                        with open(log_file_path, 'a') as log_file:
                            log_file.write(f"Error occurred for product code {product_code}: {e}\n")
                        go_to_homepage(driver)
                        continue

            driver.quit()
            logger.info("Process completed.")
            break  # 작업이 성공적으로 완료되면 반복을 중단
        except Exception as e:
            logger.error(f"An unexpected error occurred: {e}")
            logger.info("Restarting in 5 seconds...")
            with open(log_file_path, 'a') as log_file:
                log_file.write(f"Unexpected error: {e}\n")
            time.sleep(5)

if __name__ == "__main__":
    main()
