import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1. 깃허브 시크릿(Secrets) 정보 가져오기
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 경로 설정 (파일을 임시로 받을 폴더 생성)
current_dir = os.getcwd()
temp_download_dir = os.path.join(current_dir, 'temp_downloads')
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 설정 (화면 없는 Headless 모드 + 보안 무시)
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--safebrowsing-disable-download-protection')

options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
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
    print("조회 기간 설정 및 데이터를 조회합니다...")
    s_date_input = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(15) 

    # 7. 엑셀 다운로드 클릭
    print("엑셀 다운로드를 시작합니다...")
    excel_btn = driver.find_element(By.CLASS_NAME, 'excel_btn')
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 다운로드 완료를 위한 충분한 대기 시간
    print("다운로드 완료를 기다리는 중 (30초)...")
    time.sleep(30) 

    # 8. 파일명 변경 및 덮어씌우기
    files = os.listdir(temp_download_dir)
    if files:
        # 가장 최근 파일 찾기 (임시 파일 제외)
        actual_files = [f for f in files if not f.endswith('.crdownload')]
        actual_files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
        latest_file = actual_files[-1]
        
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        # 기존 파일이 있다면 삭제 후 교체
        if os.path.exists(target_path):
            os.remove(target_path)
            
        shutil.move(source_path, target_path)
        print(f"✅ 성공: current_sales.xlsx 파일이 업데이트되었습니다.")
    else:
        print("❌ 실패: 다운로드 폴더에 파일이 존재하지 않습니다.")

finally:
    driver.quit()