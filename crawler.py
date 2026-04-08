import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# 크롬 실행
driver = webdriver.Chrome()
driver.get("https://energyinfo.seoul.go.kr/energy/energyUsagePattern?menu-id=Z020400")

wait = WebDriverWait(driver, 15)

years = [str(y) for y in range(2025, 2017, -1)]

for year in years:
    print(f"{year} 시작")

    # 1. 다운로드 버튼 클릭 (팝업 열기)
    download_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "downloadButton"))
    )
    driver.execute_script("arguments[0].click();", download_btn)

    # 2. 팝업 내부 select 로딩 기다리기
    wait.until(EC.presence_of_element_located((By.ID, "downYear")))

    # 3. 옵션 선택
    Select(driver.find_element(By.ID, "downYear")).select_by_value(year)
    Select(driver.find_element(By.ID, "downSigCd")).select_by_value("")        # 전체
    Select(driver.find_element(By.ID, "downType")).select_by_value("POWER")    # 전기
    Select(driver.find_element(By.ID, "downDong")).select_by_value("hjDong")   # 행정동

    #  중요: 데이터 로딩 기다리기
    time.sleep(2)

    # 4. 엑셀 다운로드 (JS 직접 실행)
    driver.execute_script("downExcel();")

    print(f"{year} 다운로드 완료")

    # 5. alert 뜨면 처리
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

    # 6. 다운로드 시간 대기
    time.sleep(3)

    # 7. 팝업 닫기
    try:
        close_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//dialog//button"))
        )
        close_btn.click()
    except:
        pass

    time.sleep(1)

driver.quit()