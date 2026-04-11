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
current_dir = os.path.abspath(os.getcwd())
temp_download_dir = os.path.join(current_dir, 'temp_downloads')

if os.path.exists(temp_download_dir):
    shutil.rmtree(temp_download_dir)
os.makedirs(temp_download_dir, exist_ok=True)

# 3. 브라우저 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument('--headless=new') # 최신 헤드리스 모드 사용
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--allow-running-insecure-content')

# 다운로드 자동 승인 설정
options.add_experimental_option("prefs", {
    "download.default_directory": temp_download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 💡 중요: 헤드리스 모드에서 다운로드를 강제 허용하는 명령 (보안 경고 무시)
driver.execute_cdp_cmd('Page.setDownloadBehavior', {
    'behavior': 'allow',
    'downloadPath': temp_download_dir
})

try:
    # 4. 로그인
    print("🚀 WOS 접속 및 로그인 시작...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    # 5. 매출현황 이동
    print("📂 매출현황 페이지 이동 중...")
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)

    # 6. 날짜 조회
    print("🔍 데이터 조회 버튼 클릭 (2026-01-01 기준)...")
    s_date_input = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date_input)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(25) 

    # 7. 엑셀 다운로드 클릭
    print("📥 엑셀 다운로드 버튼 클릭 시도...")
    excel_btn = driver.find_element(By.CLASS_NAME, 'excel_btn')
    driver.execute_script("arguments[0].click();", excel_btn)
    
    # 8. 파일 생성 확인
    print("⏳ 파일 생성 대기 중 (보안 경고 자동 통과 시도)...")
    latest_file = None
    for i in range(50):
        # .crdownload가 없는 순수한 파일이 생길 때까지 대기
        files = [f for f in os.listdir(temp_download_dir) 
                 if not f.endswith('.crdownload') and not f.startswith('.')]
        
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(temp_download_dir, x)))
            latest_file = files[-1]
            print(f"✅ 다운로드 완료: {latest_file}")
            break
        
        if i % 5 == 0:
            print(f"... 현재 폴더 상태: {os.listdir(temp_download_dir)}")
        time.sleep(2)

    # 9. 파일 변환 및 저장
    if latest_file:
        source_path = os.path.join(temp_download_dir, latest_file)
        target_path = os.path.join(current_dir, "current_sales.xlsx")
        
        print(f"📊 엑셀 변환 중 (Pandas)...")
        try:
            # WOS 특유의 HTML 형식을 읽어 진짜 엑셀(.xlsx)로 저장
            df_list = pd.read_html(source_path, flavor='html5lib')
            if df_list:
                df = df_list[0]
                df.to_excel(target_path, index=False)
                print(f"🎉 성공적으로 완료되었습니다!")
        except Exception as e:
            print(f"⚠️ 변환 실패, 강제 복사: {e}")
            shutil.copy2(source_path, target_path)
    else:
        raise Exception("❌ 결국 파일을 받지 못했습니다. 사이트 반응이 너무 느립니다.")

except Exception as e:
    print(f"❌ 에러 발생: {e}")
    exit(1)
finally:
    driver.quit()
