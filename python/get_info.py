from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time

# 사전 정보
login_url = 'https://sso.opsnow.com/loginForm.do'
bill_url = 'https://metering.opsnow.com/billing'
login_id = ''
login_pw = ''



try:

    # Options
    options = webdriver.ChromeOptions()
    options.add_argument('window-size=1920x1080')
    options.add_argument('disable-gpu')

    # chrome driver 실행
    driver = webdriver.Chrome(executable_path='chromedriver.exe')

    # Login 페이지로 이동
    driver.get(login_url)
    driver.implicitly_wait(time_to_wait=5)

    # 계정 정보 기입 후 로그인
    id_box = driver.find_element_by_id("username")
    id_box.send_keys(login_id)
    pw_box = driver.find_element_by_id("password")
    pw_box.send_keys(login_pw)
    driver.find_element_by_xpath('//*[@id="loginForm"]/p[6]/button').click()
    driver.implicitly_wait(time_to_wait=5)

    # 1. Billing 페이지로 이동
    driver.get(bill_url)
    time.sleep(3)

    # 2. 고객사 List 뽑기
    # html = driver.page_source
    # soup = BeautifulSoup(html, 'html.parser')
    # customers = soup.find("ul", "list-companies")
    # cust_list = customers.getText().split("\n\n")
    # cust_list = list(filter(None, cust_list))
    # print(len(cust_list))

    elem = driver.find_elements_by_xpath('//*[@id="select-company"]/div/ul')
    print(len(elem))
    for i in elem:
        try:
            driver.find_element_by_xpath('//*[@id="select-company"]/button').click()
            time.sleep(2)
            i.click()
            time.sleep(2)
        except:
            pass
    time.sleep(2)





except Exception as e:
    print("Error 발생", e)

finally:
    driver.quit()
