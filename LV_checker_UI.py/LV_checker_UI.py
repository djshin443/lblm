import requests
import json
import tkinter as tk
from tkinter import scrolledtext, messagebox
from threading import Thread
from datetime import datetime
import time
import schedule
import random


def get_stores(textbox):
    skuId = input_entry.get().upper()  # Convert SKU ID to uppercase
    if not skuId:
        messagebox.showwarning("Warning", "Please enter a SKU ID")
        return

    # Update status
    status_label.config(text="Checking stores...")
    root.update_idletasks()

    url = 'https://api.louisvuitton.com/eco-eu/search-merch-eapi/v1/kor-kr/stores/query'

    # Chrome-like headers with all required fields
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
        'client_secret': '60bbcdcD722D411B88cBb72C8246a22F',
        'client_id': '607e3016889f431fb8020693311016c9',
        'Origin': 'https://kr.louisvuitton.com',
        'Referer': 'https://kr.louisvuitton.com/kor-kr/products/pochette-voyage-souple-monogram-other-nvprod6210015v/M13962',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'sec-ch-ua': '"Chromium";v="134", "Not:A-Brand";v="24", "Google Chrome";v="134"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-site',
        # Add more browser-like headers
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Priority': 'u=1, i',
        'Connection': 'keep-alive'
    }

    # Random session ID to simulate browser behavior
    session_id = ''.join(random.choice('0123456789abcdefABCDEF') for _ in range(32))
    timestamp = str(int(time.time()))

    # Detailed cookie string like browser would send
    cookies = {
        'AKA_A2': 'A',
        'lv-dispatch': 'kor-kr',
        'lv-dispatch-url': 'https://kr.louisvuitton.com/kor-kr/products/pochette-voyage-souple-monogram-other-nvprod6210015v/M13962',
        'ak_cc': 'KR',
        'ak_bmsc': f'session_id_{session_id}~{timestamp}',
        'OPTOUTMULTI': '0:0%7Cc1:0%7Cc2:0%7Cc4:0%7Cc3:0%7Cc5:0',
        '_ga': f'GA1.1.{random.randint(100000000, 999999999)}.{timestamp}',
        '_ga_S6ED35NYJQ': f'GS1.1.{timestamp}.1.0.{timestamp}.0.0.0'
    }

    data = {
        "flagShip": False,
        "country": "KR",
        "query": "",
        "clickAndCollect": False,
        "skuId": skuId,
        "pageType": "productsheet"
    }

    try:
        # Create a session to maintain cookies
        session = requests.Session()

        # Set the user agent and other persistent headers
        session.headers.update(headers)

        # Update cookies
        for key, value in cookies.items():
            session.cookies.set(key, value)

        # Make the request with the session
        response = session.post(url, json=data)

        print("Response status code:", response.status_code)
        print("Response headers:", response.headers)
        print("Response content snippet:", response.text[:500])  # Print first 500 chars of response content

        if response.status_code == 200:
            textbox.config(state='normal')  # enable editing

            # Clear previous results
            textbox.delete(1.0, tk.END)

            # Parse JSON response
            response_json = response.json()
            stores = response_json.get('hits', [])

            if stores:
                textbox.insert(tk.INSERT, f"=== Results for SKU: {skuId} ===\n\n")
                available_stores = False

                for store in stores:
                    additional_properties = store.get('additionalProperty', [])
                    is_available = False

                    # Check if the item is available in this store
                    for item in additional_properties:
                        if item.get('name') == 'stockAvailability' and item.get('value') == 'true':
                            is_available = True
                            available_stores = True
                            break

                    if is_available:
                        textbox.insert(tk.INSERT, f"✓ {store.get('name')}\n")

                        # Add address if available
                        address = store.get('address', {})
                        if address:
                            street_address = address.get('streetAddress', '')
                            locality = address.get('addressLocality', '')
                            region = address.get('addressRegion', '')
                            postal_code = address.get('postalCode', '')

                            address_parts = []
                            if street_address:
                                address_parts.append(street_address)
                            if locality:
                                address_parts.append(locality)
                            if region:
                                address_parts.append(region)
                            if postal_code:
                                address_parts.append(postal_code)

                            if address_parts:
                                textbox.insert(tk.INSERT, f"   주소: {', '.join(address_parts)}\n")

                        # Add telephone if available
                        telephone = store.get('telephone', '')
                        if telephone:
                            textbox.insert(tk.INSERT, f"   전화: {telephone}\n")

                        # Add a blank line after each store for better readability
                        textbox.insert(tk.INSERT, "\n")

                if not available_stores:
                    textbox.insert(tk.INSERT, '⚠️ 아이템이 어떤 매장에서도 이용할 수 없습니다.\n')
            else:
                textbox.insert(tk.INSERT, '⚠️ 매장 정보를 찾을 수 없습니다.\n')

            textbox.insert(tk.INSERT, "\n" + "-" * 50 + "\n")
            textbox.insert(tk.INSERT, "마지막 업데이트: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
            textbox.yview(tk.END)  # Scroll the textbox to the end
            textbox.config(state='disabled')  # disable editing

            # Update status
            status_label.config(
                text=f"완료. {len([s for s in stores if any(item.get('name') == 'stockAvailability' and item.get('value') == 'true' for item in s.get('additionalProperty', []))])}개 매장에서 이용 가능")
        else:
            print("Error occurred:", response.status_code)
            textbox.config(state='normal')
            textbox.delete(1.0, tk.END)
            textbox.insert(tk.INSERT, f"⚠️ 오류: HTTP 상태 코드 {response.status_code}\n")
            textbox.insert(tk.INSERT, "마지막 업데이트: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
            textbox.config(state='disabled')
            status_label.config(text=f"오류: HTTP 상태 코드 {response.status_code}")
    except Exception as e:
        print(f"Exception occurred: {e}")
        textbox.config(state='normal')
        textbox.delete(1.0, tk.END)
        textbox.insert(tk.INSERT, f"⚠️ 오류: {e}\n")
        textbox.insert(tk.INSERT, "마지막 업데이트: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        textbox.config(state='disabled')
        status_label.config(text=f"오류: {str(e)[:50]}")


def update_schedule():
    # Clear all existing jobs
    schedule.clear()

    skuId = input_entry.get()
    update_freq = freq_var.get()

    if not skuId:
        messagebox.showwarning("Warning", "Please enter a SKU ID")
        return

    # First run immediately
    get_stores(st)

    if update_freq != "once":
        # Schedule the job according to the selected frequency
        if update_freq == "hourly":
            schedule.every().hour.do(get_stores, textbox=st)
            status_label.config(text="자동 업데이트: 매시간")
        elif update_freq == "every_5_min":
            schedule.every(5).minutes.do(get_stores, textbox=st)
            status_label.config(text="자동 업데이트: 5분마다")
        elif update_freq == "every_30_sec":
            schedule.every(30).seconds.do(get_stores, textbox=st)
            status_label.config(text="자동 업데이트: 30초마다")


def on_exit():
    if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
        root.destroy()


# Set up the UI
root = tk.Tk()
root.title('LV 매장 재고 확인')
root.geometry('550x450')
root.protocol("WM_DELETE_WINDOW", on_exit)

# Create a frame for the input area
input_frame = tk.Frame(root, pady=10)
input_frame.pack(fill='x')

input_label = tk.Label(input_frame, text='SKU 번호:')
input_label.pack(side='left', padx=5)

input_entry = tk.Entry(input_frame, width=20)
input_entry.pack(side='left', padx=5)
input_entry.focus()  # Set focus to the entry field

# Add more update frequency options
freq_var = tk.StringVar(root)
freq_var.set("once")  # default value
freq_label = tk.Label(input_frame, text='업데이트:')
freq_label.pack(side='left', padx=5)
freq_option = tk.OptionMenu(input_frame, freq_var, "once", "every_30_sec", "every_5_min", "hourly")
freq_option.pack(side='left', padx=5)

button = tk.Button(input_frame, text='확인', command=update_schedule)
button.pack(side='left', padx=10)

# Help text
help_frame = tk.Frame(root)
help_frame.pack(fill='x', padx=10)
help_text = tk.Label(help_frame, text="입력 예: M13962 (대소문자 구분 없음)", fg="gray")
help_text.pack(anchor='w')

# Status area - scrollable text box
st = scrolledtext.ScrolledText(root, height=20, state='disabled')
st.pack(expand=True, fill='both', padx=10, pady=10)

# Add status label at the bottom
status_label = tk.Label(root, text="준비 완료", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(side=tk.BOTTOM, fill=tk.X)


# Define a function for running the scheduler in a loop
def run_scheduler():
    while True:
        try:
            schedule.run_pending()
        except Exception as e:
            print(f"Scheduler error: {e}")
        time.sleep(1)


# Run the scheduler in a separate thread
scheduler_thread = Thread(target=run_scheduler)
scheduler_thread.daemon = True  # Set the thread as a daemon so it terminates when the main program exits
scheduler_thread.start()


# Handle Enter key to trigger check
def on_enter(event):
    update_schedule()


root.bind('<Return>', on_enter)

root.mainloop()
