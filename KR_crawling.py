from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
import openpyxl 
from datetime import datetime
import sys
import os.path



rating_period = input("기간을 설정하세요(당일공시/최근1주일/최근2주일/최근1개월): ")
excel_name = input("엑셀파일명: ")


wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(["SPC","회차","증권종류","발행금액","발행일","만기일","평가종류","평가일","공시일","등급"])
now = datetime.now()
date = now.strftime('%Y.%m.%d')
excel_filename = excel_name +"_" + rating_period +"_" + date + "(KR).xlsx"

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('headless')
chrome_options.add_argument('window-size=1920x1080')
chrome_options.add_argument("disable-gpu")

# Chromedriver를 실행파일에 포함시키기 위한 코드
if getattr(sys, 'frozen', False):
    chromedriver = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver, options=chrome_options)
else:
    chromedriver = './chromedriver'
    driver = webdriver.Chrome(chromedriver, options=chrome_options)

driver.implicitly_wait(1)
driver.get('http://www.rating.co.kr/disclosure/QDisclosure002.do')

select=Select(driver.find_element_by_id("evalDt"))
select.select_by_visible_text(rating_period)
select=Select(driver.find_element_by_id("svcty"))
select.select_by_visible_text("ABS")

button = driver.find_element_by_class_name("btn_small")
button.click()

table = driver.find_element(By.XPATH, "/html/body/div/div/div/div[3]/div[2]/div/div/table/tbody")

trs = table.find_elements(By.TAG_NAME, "tr")
count_trs = len(trs)
j=1
for i in range(count_trs):
    table = driver.find_element(By.XPATH, "/html/body/div/div/div/div[3]/div[2]/div/div/table/tbody")
    table.find_element(By.XPATH, "/html/body/div/div/div/div[3]/div[2]/div/div/table/tbody/tr[{}]/td[1]/a".format(j)).click()
    driver.find_element(By.XPATH, "/html/body/div[2]/div/form[1]/div/div[3]/div[1]/ul/li[2]/h3/a").click()
    detail_spc = driver.find_element(By.XPATH, "/html/body/div[2]/div/form[1]/div/div[2]/div[2]/div/table/tbody/tr[1]/td[1]")
    detail_table = driver.find_element(By.XPATH, "/html/body/div[2]/div/form[1]/div/div[3]/div[3]/div[1]/div[2]/div/div/table/tbody")
    for detailtr in detail_table.find_elements(By.TAG_NAME, "tr"):
        td = detailtr.find_elements(By.TAG_NAME, "td")
        s = detail_spc.text
        s0 = td[0].text
        s1 = td[1].text
        s2 = td[2].text
        s3 = td[3].text
        s4 = td[4].text
        s5 = td[5].text
        s6 = td[6].text
        s7 = td[7].text
        s8 = td[8].text
        sheet.append([s,s0,s1,s2,s3,s4,s5,s6,s7,s8])
    j = j + 1
    driver.back()

wb.save(excel_filename)
time.sleep(2)
