from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pyautogui
import time
import pandas as pd
from selenium.webdriver.common.keys import Keys
import keyboard
import win32api
import shutil









driver = webdriver.Chrome("C:\Program Files (x86)\chromedriver_win32\chromedriver.exe")
driver.get("https://www.office.com/launch/forms?auth=2")
win32api.LoadKeyboardLayout('00000409',1) # to switch to english
#initialize operations with excel
ExcelOperator = ExcelDownloader(driver=driver)
#initialize parameters for logging into outlook
LoginOperator = LoginManager(driver=driver)

time.sleep(6)
#maximize window and log in
LocateMoveAndClick("Box2.png",confidence=0.8)
time.sleep(4)
LoginOperator.login()
time.sleep(2)
result_list = ExcelOperator.scrape_excel_blocks()

for link,label in result_list:
    if "Math Commitee" in label:
        final_label = label
        final_link = link
        break

#add a new file
print(final_link)
ExcelOperator.add_new_page(final_link)
time.sleep(2)
    
#click and download the excel file containing all the responses
LocateMoveAndClick("ExcelButton.png",confidence=0.9)

#wait until I close the program
keyboard.wait("esc")

#delete all cokies and close the browser
driver.delete_all_cookies()
driver.close()






