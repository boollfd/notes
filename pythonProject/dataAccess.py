# -*-coding = utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


import os
import time


chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:61124")
chrome_driver = r"C:\Users\Administrator\Desktop\tools\toolsPackage\chromedriver.exe"
driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)

# all_windows = driver.window_handles
# print(all_windows)
# for handle in all_windows:
#     driver.switch_to.window(handle)
#     print(driver.title)
#     print(driver.current_url)
driver.get("https://admin.pupumall.com/statistics/stroeProductIncomeTaskNew/Index")
element = WebDriverWait(driver, 5, 0.5).until(EC.presence_of_element_located((By.ID, "content-container")))
# buttons = element.find_elements_by_class_name("ant-btn")
# print(len(buttons))
# buttons[2].click()

# dateBegin = "2021-02-20"
# dateEnd = "2021-02-22"
# city = ["福州市"]

# element = WebDriverWait(driver, 5, 0.5).until(EC.presence_of_element_located((By.ID, "content-container")))
# rows = element.find_elements_by_class_name("ant-row")

# dates = rows[2].find_elements_by_tag_name("input")
# rows[2].click()
# calendarPanel = WebDriverWait(driver, 5, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME, "ant-calendar-panel")))
# dates = calendarPanel.find_elements_by_tag_name("input")
# dates[0].click()
# dates[0].send_keys(Keys.CONTROL, 'a')
# dates[0].send_keys(dateBegin)
# dates[1].click()
# dates[1].send_keys(Keys.CONTROL, 'a')
# dates[1].send_keys(dateEnd)
# driver.find_element_by_class_name("ant-card-head-title").click()

# rows[3].find_elements_by_tag_name("input")[1].click()
# citys = WebDriverWait(driver, 5, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME, "ant-checkbox-group")))
# citys = citys.find_elements_by_tag_name("label")
# for i in citys:
#     if i.text in city:
#         i.click()

# rows[4].click()


Authorization=""
cookies = driver.get_cookies()
print(cookies)
for cookie in cookies:
    if cookie['name'] == "Authorization":
        Authorization = cookie['value']


os.system('taskkill /im chromedriver.exe /F')
