from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import shutil
import pyautogui
import time
import keyboard
import win32api
import pandas as pd
import os
from selenium.webdriver.common.keys import Keys
from objects.ExcelDownloader import ExcelDownloader
from objects.ExcelManager import ExcelManager
from objects.LoginManager import LoginManager,LeaderboardManager

if __name__ == "__main__":
    # driver = webdriver.Chrome("C:\Program Files (x86)\chromedriver_win32\chromedriver.exe")

    #initialize objects
    # Logins = LoginManager(driver)
    # ExcelDownloading = ExcelDownloader(driver)
    # ExcelManagement = ExcelManager(driver)
    basepath = os.getcwd()

    driver = webdriver.Chrome("C:\Program Files (x86)\chromedriver_win32\chromedriver.exe")
    driver.get("https://www.office.com/launch/forms?auth=2")
    win32api.LoadKeyboardLayout('00000409',1) # to switch to english
    #initialize operations with excel
    ExcelOperator = ExcelDownloader(driver=driver)
    #initialize parameters for logging into outlook
    LoginOperator = LoginManager(driver=driver)

    time.sleep(6)
    #maximize window and log in
    ExcelOperator.LocateMoveAndClick(str(os.path.join(basepath,"Box2.png")),confidence=0.8)
    time.sleep(4)
    LoginOperator.login()
    time.sleep(4)
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
    ExcelOperator.LocateMoveAndClick(str(os.path.join(basepath,"ExcelButton.png")),confidence=0.9)

    #wait until I close the program
    # keyboard.wait("esc")


    #move newly downloaded file
    time.sleep(4)
    weekly_scores = ExcelOperator.find_newly_downloaded_files(final_label)
    shutil.move(weekly_scores,basepath)

    #initialize the excel leaderboard object which will scrape the leaderboard table
    ExcelLeaderBoard = LeaderboardManager(driver=driver)
    

    driver.get("https://kist.learning.powerschool.com/yuqi.zhao/mathcommittee/wk/13661263/wiki/view/62652247#/")
    time.sleep(6)
    #maximize window and log in
    ExcelLeaderBoard.signin(str(os.path.join(basepath,"Box2.png")),confidence=0.8)
    time.sleep(4)
    ExcelLeaderBoard.signin(str(os.path.join(basepath,"Signin.png")),confidence=0.9)
    time.sleep(2)   
    LoginOperator.login()
    time.sleep(2)

    #save the scraped table
    ac = LeaderboardManager.read_table_and_save_to_excel()
    df = LeaderboardManager.clean_leaderboard_table(ac)
    LeaderboardManager.save_results(df,os.path.join(os.getcwd(),"AnswerDetection","Math Commitee Leaderboard.xlsx"))

    weekly_leaderboard = ExcelManager.find_newly_downloaded_files()
    ExcelManagement = ExcelManager.read_excel_file(path=str(weekly_scores,weekly_leaderboard))
    updated = ExcelManagement.update_leaderboard()
    
    LeaderboardManager.save_results(updated,os.path.join(os.getcwd(),"AnswerDetection","Math Commitee Leaderboard New.xlsx"))
    driver.close()



    
    






