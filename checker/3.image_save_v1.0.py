import os
import re
import requests
import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Setup logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Additional logger for image download info
download_logger = logging.getLogger('download_logger')
download_logger.setLevel(logging.INFO)
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
download_logger.addHandler(console_handler)

def setup_driver():
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument('--disable-infobars')
    #options.add_argument('--headless')
    options.add_argument('--start-maximized')  # 창 최대화
    options.add_argument('--log-level=3')
    options.add_argument('--silent')  # Suppress all console output from Chrome
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def download_image(image_url, folder, filename):
    if not os.path.exists(folder):
        os.makedirs(folder)

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
    }

    try:
        response = requests.get(image_url, headers=headers)
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
    except requests.exceptions.RequestException as e:
        download_logger.error(f"Failed to download {image_url}")
        download_logger.error(f"Error: {e}")

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
        search_button.click()
        time.sleep(2)

        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#searchHeaderInput"))
        )
        search_input.clear()
        search_input.send_keys(product_code)
        search_input.send_keys(Keys.ENTER)
        time.sleep(5)

        # 팝업 닫기
        close_popup(driver)
    except Exception as e:
        logger.error(f"검색 중 오류 발생: {e}")


def close_popup(driver):
    try:
        close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".lv-notifications__close.lv-button.-only-icon"))
        )
        close_button.click()
        logger.info("Popup closed")
    except Exception as e:
        logger.info("No popup found to close")

def get_image_urls(driver, product_code):
    logger.info(f"Searching images for product code: {product_code}")
    search_product(driver, product_code)

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
            driver.get("https://kr.louisvuitton.com/")

            for sheet_name, product_codes in all_product_codes.items():
                sanitized_sheet_name = sanitize_folder_name(sheet_name)
                for product_code in product_codes:
                    sanitized_product_code = sanitize_folder_name(str(product_code))
                    try:
                        image_urls = get_image_urls(driver, product_code)
                        if not image_urls:
                            continue

                        folder_path = os.path.join("image", sanitized_sheet_name, sanitized_product_code)

                        for idx, url in enumerate(image_urls):
                            logger.info(f"Downloading image: {url}")
                            download_image(url, folder_path, f"{sanitized_product_code}_image_{idx + 1}")
                    except Exception as e:
                        logger.error(f"Error occurred while processing product code {product_code}: {e}")
                        continue

            driver.quit()
            logger.info("Process completed.")
            break  # 작업이 성공적으로 완료되면 반복을 중단
        except Exception as e:
            logger.error(f"An unexpected error occurred: {e}")
            logger.info("Restarting in 5 seconds...")
            time.sleep(5)

if __name__ == "__main__":
    main()
