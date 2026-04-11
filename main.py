import os, time, shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# 1. 시크릿 정보
USER_ID = os.environ.get('WOS_ID')
USER_PW = os.environ.get('WOS_PW')
current_dir = os.path.abspath(os.getcwd())
temp_dir = os.path.join(current_dir, 'temp_downloads')
os.makedirs(temp_dir, exist_ok=True)

# 2. 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_experimental_option("prefs", {"download.default_directory": temp_dir})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.execute_cdp_cmd('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': temp_dir})

try:
    print("🚀 WOS 실적 수집 시작...")
    driver.get('http://wos.bridgestone-korea.co.kr/')
    time.sleep(5)
    driver.find_element(By.ID, 'userID').send_keys(USER_ID)
    driver.find_element(By.ID, 'userPIN').send_keys(USER_PW)
    driver.find_element(By.CLASS_NAME, 'login_btn').click()
    time.sleep(10) 

    # 매출현황 메뉴 이동
    driver.get('http://wos.bridgestone-korea.co.kr/inqireMgmt/selngLedgrInqire2.do')
    time.sleep(10)
    
    # 2026년 1월 1일부터 조회
    s_date = driver.find_element(By.ID, 'sDate')
    driver.execute_script("arguments[0].value = '2026-01-01';", s_date)
    driver.find_element(By.CLASS_NAME, 'search_btn').click()
    time.sleep(25) 
    driver.find_element(By.CLASS_NAME, 'excel_btn').click()
    time.sleep(30)

    # 3. 💡 [LT 비법] 데이터 합성 (Synthesis)
    files = [f for f in os.listdir(temp_dir) if not f.endswith('.crdownload')]
    if files:
        source_path = os.path.join(temp_dir, files[0])
        ref_path = os.path.join(current_dir, "historical_data.xlsx") # 뼈대
        target_path = os.path.join(current_dir, "current_sales.xlsx") # 완성본

        # WOS 데이터 읽기 (HTML/CP949)
        with open(source_path, 'r', encoding='cp949', errors='ignore') as f:
            new_df = pd.read_html(f, flavor='html5lib')[0]
        
        # 최신 실적을 거래처/사이즈/패턴별로 합산 (LT와 동일 로직)
        new_agg = new_df.groupby(['거래처', 'SIZE', 'PTTN'])['합계수량'].sum().reset_index()

        # 기준표(Historical) 읽기
        ref_df = pd.read_excel(ref_path)

        # 💡 매핑 엔진: 거래처+사이즈+패턴 세 가지가 모두 맞을 때만 숫자를 넣습니다.
        def get_actual_sales(row):
            match = new_agg[
                (new_agg['거래처'].astype(str).str.strip() == str(row['거래처명']).strip()) &
                (new_agg['SIZE'].astype(str).str.strip() == str(row['사이즈']).strip()) &
                (new_agg['PTTN'].astype(str).str.strip() == str(row['패턴명']).strip())
            ]
            return match['합계수량'].sum() if not match.empty else 0

        print("📊 실적 데이터 매핑 중...")
        ref_df['2026년(당해)'] = ref_df.apply(get_actual_sales, axis=1)
        
        # 최종 완성본 저장
        ref_df.to_excel(target_path, index=False)
        print("🎉 합성 완료! current_sales.xlsx가 생성되었습니다.")

finally:
    driver.quit()
    shutil.rmtree(temp_dir)
