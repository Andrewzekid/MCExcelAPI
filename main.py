from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pyautogui
import time
from selenium.webdriver.common.keys import Keys


def signin(path,confidence=0.9):
    position_of_signin = pyautogui.locateCenterOnScreen(path,confidence=confidence)
    pyautogui.moveTo(position_of_signin[0],position_of_signin[1],duration=2)
    pyautogui.click(button="left")
    time.sleep(0.2)

#code to login
def login():
    try:
        element = WebDriverWait(driver,50).until(
            EC.presence_of_element_located((By.XPATH,'//input[@name="loginfmt"]'))
        )
    except:
        pass
    else:
        #log into email adress and acsess inbox
        driver.find_element_by_xpath('//input[@name="loginfmt"]').send_keys("yw3479@email.kist.ed.jp"+Keys.ENTER)
        driver.implicitly_wait(10)
        ITbutton = driver.find_element_by_xpath('//div[@id="aadTile"]')
        webdriver.ActionChains(driver).click(ITbutton).perform()
        webdriver.ActionChains(driver).release().perform()
        driver.implicitly_wait(5)
        driver.find_element_by_xpath('//input[@name="passwd"]').send_keys("Nut5flower#"+Keys.ENTER)
        driver.implicitly_wait(5)
        driver.find_element_by_xpath('//input[@id="idSIButton9"]').send_keys(Keys.ENTER)

driver = webdriver.Chrome("C:\Program Files (x86)\chromedriver_win32\chromedriver.exe")
driver.get("https://kist.learning.powerschool.com/yuqi.zhao/mathcommittee/wk/13661263/wiki/view#/")
time.sleep(6)
#maximize window and log in
signin("Box2.png",confidence=0.8)
time.sleep(4)
signin("Signin.png",confidence=0.9)
time.sleep(2)   
login()




