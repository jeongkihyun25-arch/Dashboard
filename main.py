import os
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1. 시크릿 정보
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 경로 설정
current_dir = os.getcwd()
temp_download_dir = os.path.join(current_dir, 'temp_downloads')

if os.path.exists(temp_download_dir):
    shutil.rmtree(temp_download_dir)
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')

options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": False,
    "profile.default_content_setting_values.automatic_downloads": 1
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    # 4. 로그인
    print("WOS 접속 중...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    # 5. 매출현황 이동
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)

    # 6. 날짜 조회
    s_date_input = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    print("데이터 조회 중...")
    time.sleep(20) 

    # 7. 엑셀 다운로드
    excel_btn = driver.find_element(By.CLASS_NAME, 'excel_btn')
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 8. 다운로드 대기 및 진짜 파일 찾기
    print("파일 다운로드 확인 중...")
    latest_file = None
    for _ in range(30):
        # 💡 중요: .crdownload(다운로드중) 이거나 . (숨김임시파일) 로 시작하는 파일은 제외
        files = [f for f in os.listdir(temp_download_dir) 
                 if not f.endswith('.crdownload') and not f.startswith('.')]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            break
        time.sleep(2)

    # 9. 파일 변환
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        print(f"변환 시작: {latest_file} -> current_sales.xlsx")
        try:
            # 💡 html5lib을 사용하여 HTML 표를 읽음
            df_list = pd.read_html(source_path, flavor='html5lib')
            if df_list:
                df = df_list[0]
                df.to_excel(target_path, index=False)
                print(f"✅ 변환 성공!")
        except Exception as e:
            print(f"⚠️ 변환 실패, 원본 유지 시도: {e}")
            shutil.move(source_path, target_path)
    else:
        raise Exception("❌ 다운로드된 파일을 찾지 못했습니다.")

finally:
    driver.quit()
