from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
import openpyxl 

wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(["1"])


chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome('C:/Youngsu/02. 상시업무/10. 자동화프로젝트/env/ratings/chromedriver', options=chrome_options)
driver.implicitly_wait(3)
driver.get('http://www.rating.co.kr/disclosure/QDisclosure002.do')

select=Select(driver.find_element_by_id("evalDt"))
select.select_by_visible_text("당일공시")
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
    detail_spc = driver.find_element(By.XPATH, "/html/body/div[2]/div/form[1]/div/div[2]/div[2]/div/table/tbody/tr[1]/td[1]")
    detail_table = driver.find_element(By.XPATH, "/html/body/div[2]/div/form[1]/div/div[3]/div[3]/div[1]/div[2]/div/div/table/tbody")
    for detailtr in detail_table.find_elements(By.TAG_NAME, "tr"):
        td = detailtr.find_elements(By.TAG_NAME, "td")
        s = "{} , {} , {} , {} , {} , {} , {} , {} , {} , {} \n".format(detail_spc.text, td[0].text, td[1].text , td[2].text, td[3].text, td[4].text, td[5].text, td[6].text, td[7].text, td[8].text)
        print(s)
        sheet.append([s])
    j = j + 1
    print(j)
    driver.back()

wb.save("today1.xlsx")

# j = 0
# print(tr_numb)
# for row in range(tr_numb):
#     trs.find_elements(By.TAG_NAME, "a")[j].click()
#     driver.back()  # 페이지 뒤로 가기
#     time.sleep(4)



# search_input = driver.find_element_by_id('evalDt')
# search_input.send_keys('케이비증권')
# search_input.send_keys(Keys.RETURN)
# print(driver.title)

# 코드 실행을 잠시 멈춘다.
time.sleep(2)

# 사용을 마치면 드라이버를 종료시킨다. 
# driver.quit()

# for tr in table.find_elements(By.TAG_NAME, "tr"):

#     td = tr.find_elements(By.TAG_NAME, "td")
#     s = "{} , {} , {} , {} , {} , {} , {} , {} , {} \n".format(td[0].text, td[1].text , td[2].text, td[3].text, td[4].text, td[5].text, td[6].text, td[7].text, td[8].text)
#     print(s)
#     detail = tr.find_element(By.TAG_NAME, "a")
#     detail.click() 
#     detail_table = driver.find_element(By.XPATH, "/html/body/div[2]/div/form[1]/div/div[3]/div[3]/div[1]/div[2]/div/div/table/tbody")
#     for detailtr in detail_table.find_elements(By.TAG_NAME, "tr"):
#         detailtd = detailtr.find_elements(By.TAG_NAME, "td")

# time.sleep(2)
# trs = table.find_elements(By.TAG_NAME, "tr")
# count_trs = len(trs)
# j=1
# for i in range(count_trs):
#     table = driver.find_element(By.XPATH, "/html/body/div/div/div/div[3]/div[2]/div/div/table/tbody")
#     table.find_element(By.XPATH, "/html/body/div/div/div/div[3]/div[2]/div/div/table/tbody/tr[{}]/td[1]/a".format(j)).click()
#     j = j + 1
#     print(j)
#     driver.back()
