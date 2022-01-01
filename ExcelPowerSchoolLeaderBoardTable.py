from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pyautogui
import pathlib
from pathlib import Path
import os
import time
import pandas as pd
from selenium.webdriver.common.keys import Keys
p = Path(".")
base_path  = p.resolve() / "AnswerDetection"

def signin(path,confidence=0.9):
    position_of_signin = pyautogui.locateCenterOnScreen(path,confidence=confidence)
    pyautogui.moveTo(position_of_signin[0],position_of_signin[1],duration=2)
    pyautogui.click(button="left")
    time.sleep(0.2)

#code to login
def login() -> None:
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

def read_table_and_save_to_excel():
    table = driver.find_element_by_xpath('//div[@id="bbody203007212"]').get_attribute("innerHTML")
    df = pd.read_html(table)
#     file_name = "MathCommitee.xlsx"
#     df.to_excel(file_name)


    return df

def clean_leaderboard_table(ls):
    #clean the table that is scraped by pandas
    df = ls[0]
    new_header = df.iloc[0]
    df = df[1:]
    
    df = df.reset_index()
    df = df.drop(["index",0],axis=1)
    
    df.columns = new_header[1:]
    df = df.set_index("Name")
    df = df.rename(columns={"Points":"Score"})
    return df

def save_results(df,name):
    df.to_excel(name)
    

driver = webdriver.Chrome("C:\Program Files (x86)\chromedriver_win32\chromedriver.exe")
driver.get("https://kist.learning.powerschool.com/yuqi.zhao/mathcommittee/wk/13661263/wiki/view/62652247#/")
time.sleep(6)
#maximize window and log in
signin(str(base_path / "Box2.png"),confidence=0.8)
time.sleep(4)
signin(str(base_path /"Signin.png"),confidence=0.9)
time.sleep(2)   
login()
time.sleep(2)
ac = read_table_and_save_to_excel()
df = clean_leaderboard_table(ac)
save_results(df,os.path.join(os.getcwd(),"AnswerDetection","Math Commitee Leaderboard.xlsx"))


