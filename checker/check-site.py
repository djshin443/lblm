import os
import sys
import json
import requests  
from flask import Flask, request, render_template, jsonify, Response
from datetime import datetime, timedelta

# 스크립트 경로 설정
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__)

# IP 차단 목록 및 설정
blocked_ips = {}
MAX_ATTEMPTS = 5
BLOCK_DURATION = timedelta(minutes=5)
ip_failed_attempts = {}

@app.before_request
def before_request():
    ip_address = request.headers.get("X-Forwarded-For", request.remote_addr)

    # IP 차단된 경우 차단 메시지 출력
    if ip_address in blocked_ips and datetime.now() < blocked_ips[ip_address]:
        print(f"차단된 IP: {ip_address}")  # 차단된 IP 로깅
        return jsonify({"error": "IP blocked, try again later."}), 403
    
    # 인증 로직 추가
    auth = request.authorization
    if not auth or not check_auth(auth.username, auth.password):
        # 실패 횟수 증가
        if ip_address not in ip_failed_attempts:
            ip_failed_attempts[ip_address] = 1
        else:
            ip_failed_attempts[ip_address] += 1
        
        # 실패 횟수가 5번 이상이면 IP 차단
        if ip_failed_attempts[ip_address] >= MAX_ATTEMPTS:
            blocked_ips[ip_address] = datetime.now() + BLOCK_DURATION  # 차단 시간 설정
            return jsonify({"error": "IP blocked, try again later."}), 403
        
        print(f"인증 실패: {auth}")  # 인증 실패 로깅
        return authenticate()  # 인증 팝업 요청

def check_auth(username, password):
    return username == 'lablam' and password == '3438'

def authenticate():
    """인증 실패 시 인증 팝업을 요청하는 함수"""
    return Response(
        'You need to authenticate to access this resource.',
        401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'}
    )

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['POST'])
def login():
    ip_address = request.remote_addr

    # IP 차단 확인
    if ip_address in blocked_ips and datetime.now() < blocked_ips[ip_address]:
        return jsonify({"error": "IP blocked, try again later."}), 403

    # 비밀번호 검사
    password = request.form.get("password")
    correct_password = "3438"  # 정확한 비밀번호 확인

    if password == correct_password:
        ip_failed_attempts[ip_address] = 0  # 실패 횟수 초기화
        return jsonify({"message": "Login successful!"})

    # 실패 횟수 증가
    if ip_address not in ip_failed_attempts:
        ip_failed_attempts[ip_address] = 1
    else:
        ip_failed_attempts[ip_address] += 1

    # 실패 횟수가 5번 이상이면 IP 차단
    if ip_failed_attempts[ip_address] >= MAX_ATTEMPTS:
        blocked_ips[ip_address] = datetime.now() + BLOCK_DURATION  # 차단 시간 설정

    return jsonify({"error": "Incorrect password, try again."}), 401

@app.route('/check_availability', methods=['POST'])
def check_availability():
    sku_id = request.form.get('sku_id')

    if not sku_id:
        return render_template('result.html', message='SKU ID is required.')
    # API 요청 로직
    url = 'https://api.louisvuitton.com/eco-eu/search-merch-eapi/v1/kor-kr/stores/query'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
        'client_secret': '60bbcdcD722D411B88cBb72C8246a22F',
        'client_id': '607e3016889f431fb8020693311016c9',
        'Origin': 'https://kr.louisvuitton.com',
        'Referer': 'https://kr.louisvuitton.com/kor-kr/products/puzzle-flower-monogram-keyring-s00-nvprod4170081v/M01207',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cookie': 'ATG_SESSION_ID=E4i7k6N4uqdHA3we5jdEXtRl.front131-prd; ATGID=anonymous; anonymous_session=true; _dynSessConf=8923089313706214252; JSESSIONID=E4i7k6N4uqdHA3we5jdEXtRl.front131-prd; SGID=sb.springboot131-prd; bm_sz=C3631D7AA73F951FA5D4AD28574EDBAE~YAAQzTVDF/efuJ+IAQAAvrCnrhQzSqHopKHCkEzH4CDBpVSjt2CdcrK7xKD4Ur9r2WMDl2abOsq90ekwmhfXa7QEKxaH1Mk+tSyR4CGeG4q4OCSHFDM9pFimOHLrKfn5EA4a3tCLGxLgLpHfAfL739DE+UQnjZhg69ZeUaZzpRDauJvEKVUVGz6EGAUKGSmW5dSBtI4CCE+p42IFeiRRd/BI0wsYOdwZksOa0NYERiZ6T5u/Rru2R1t6RtpgvIqHB7ByYZsT7fFJJoNVgmWEvWJI24G6lg8qooPgIyXB2P9hGq/FvRsd170~3552308~3556914; AKA_A2=A; OPTOUTMULTI=0:0%7Cc1:0%7Cc2:0%7Cc4:0%7Cc3:0; lv-dispatch=kor-kr',
    }
    data = {
        "flagShip": False,
        "country": "KR",
        "query": "",
        "clickAndCollect": False,
        "skuId": sku_id,
        "pageType": "productsheet"
    }

    try:
        response = requests.post(url, headers=headers, json=data)

        if response.status_code == 200:
            stores = []
            for store in response.json().get('hits', []):
                if any(p.get('name') == 'stockAvailability' and p.get('value') == 'true' for p in store.get('additionalProperty', [])):
                    stores.append(store.get('name'))

            return render_template('result.html', stores_info=stores)
        else:
            return render_template('result.html', message='API 요청 실패.')
    except Exception as e:
        return render_template('result.html', message=f"API 요청 중 오류: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8443, ssl_context=('fkqpffkakd.crt', 'fkqpffkakd.pem'))
