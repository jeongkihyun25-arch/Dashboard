import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1. 깃허브 시크릿에서 로그인 정보 호출
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')

# 2. 작업 경로 설정
current_dir = os.getcwd()
temp_download_dir = os.path.join(current_dir, 'temp_downloads')

# 매번 깨끗한 상태에서 시작하도록 폴더 정리
if os.path.exists(temp_download_dir):
    shutil.rmtree(temp_download_dir)
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 보안 및 다운로드 설정 (최강화 버전)
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--allow-running-insecure-content')

# 자동 다운로드 및 팝업 차단 해제
options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "profile.default_content_setting_values.automatic_downloads": 1
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    # 4. 로그인 과정
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

    # 6. 날짜 설정 (2026-01-01부터) 및 데이터 조회
    print("데이터 조회를 시작합니다...")
    s_date_input = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    
    # 데이터 양이 많을 수 있으므로 로딩 대기
    print("데이터 로딩 중 (20초 대기)...")
    time.sleep(20) 

    # 7. 엑셀 다운로드 클릭 (강제 자바스크립트 실행)
    print("엑셀 다운로드를 시작합니다...")
    excel_btn = driver.find_element(By.CLASS_NAME, 'excel_btn')
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 8. 파일이 생성될 때까지 대기 루프 (최대 60초)
    print("파일이 생성되기를 기다리는 중...")
    latest_file = None
    for _ in range(30):
        # 임시 다운로드 중인 파일(.crdownload)은 제외하고 실제 파일만 체크
        files = [f for f in os.listdir(temp_download_dir) if not f.endswith('.crdownload')]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            break
        time.sleep(2)

    # 9. 최종 파일 처리 (.xls로 고정)
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xls")
        
        # 기존 파일이 있으면 삭제 후 교체
        if os.path.exists(target_path):
            os.remove(target_path)
            
        shutil.move(source_path, target_path)
        print(f"✅ 완료: {latest_file} -> current_sales.xls 로 저장되었습니다.")
    else:
        raise Exception("❌ 실패: 다운로드 폴더에서 파일을 찾을 수 없습니다.")

except Exception as e:
    print(f"⚠️ 에러 발생: {e}")
    exit(1)

finally:
    driver.quit()
