import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1. 시크릿 정보 (GitHub Secrets에서 아이디/비번 호출)
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 경로 설정
current_dir = os.getcwd()
temp_download_dir = os.path.join(current_dir, 'temp_downloads')
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 옵션 설정 (Headless 및 보안 해제 최강화)
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--allow-running-insecure-content')
options.add_argument('--safebrowsing-disable-download-protection')

options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False, 
    "profile.default_content_setting_values.automatic_downloads": 1 
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    # 4. 로그인 단계
    print("브릿지스톤 WOS 접속 중...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    
    print("로그인 정보를 입력합니다...")
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    # 5. 매출현황 페이지 이동
    print("매출현황 페이지로 이동합니다...")
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)

    # 6. 날짜 설정 (2026-01-01) 및 조회
    print("데이터 조회를 시작합니다 (2026-01-01 기준)...")
    s_date_input = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(15) 

    # 7. 엑셀 다운로드 (자바스크립트 클릭)
    print("엑셀 다운로드 요청 중...")
    excel_btn = driver.find_element(By.CLASS_NAME, 'excel_btn')
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 8. 파일 생성 대기 (최대 60초)
    print("파일 다운로드 대기 중...")
    latest_file = None
    for _ in range(30): 
        files = [f for f in os.listdir(temp_download_dir) if not f.endswith('.crdownload')]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            break
        time.sleep(2)

    # 9. 파일 이동 및 이름 변경 (반드시 .xls로 저장해야 함)
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        
        # WOS 사이트 특성상 실제 내용은 .xls 형식이므로 이름을 맞춰줍니다.
        target_path = os.path.join(current_dir, "current_sales.xls")
        
        if os.path.exists(target_path):
            os.remove(target_path)
            
        shutil.move(source_path, target_path)
        print(f"✅ 성공: current_sales.xls 업데이트 완료.")
    else:
        raise Exception("❌ 에러: 다운로드된 파일을 찾을 수 없습니다.")

except Exception as e:
    print(f"에러 발생: {e}")
    exit(1)

finally:
    driver.quit()
