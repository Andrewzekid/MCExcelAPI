from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pyautogui
import keyboard
import pathlib
import os
import time
import pandas as pd
from selenium.webdriver.common.keys import Keys



class ExcelDownloader:
    def __init__(self,driver,*args,**kwards):
        self.driver = driver

    def scrape_excel_blocks(self):
        try:
            element = WebDriverWait(self.driver,30).until(EC.presence_of_element_located((By.XPATH,'//a[@class="mru-list-item"]')))
        except:
            pass
        else:
            Data = self.driver.find_elements_by_class_name('mru-list-item')
            result_list = []
            for element in Data:
                link = element.get_attribute("href")
                label = element.get_attribute("aria-label")
                print(link)
                print(label)
                print("=================================")
                result_list.append((link,label))
            return result_list

    def predict_file_name(self,label):
        basename = label.split(".")[0]
        excel_name = basename.replace("Open ","")
        return excel_name

    @staticmethod
    def move_excel_file_into_current_directory() -> str:
        current_path = pathlib.Path("C:/Users/yw347/Downloads")
        answer_detection = current_path.resolve() / "AnswerDetection"
        excel_files = list(answer_detection.glob("*.xlsx"))
        if excel_files != None:
            print("Found: {}".format(str(excel_files[0])))
            return str(excel_files[0])
        else:
            raise OSError("No New Excel Download Data Detected at: {}".format(str(current_path)))

    def find_newly_downloaded_files(self,label):
        excel_name = self.predict_file_name(label)
        current_path = pathlib.Path("C:/Users/yw347/Downloads")
        weekly_scores = None
        weekly_leaderboard = None
        excel_files = list(current_path.glob("*.xlsx"))

        print(excel_files)
        for file in excel_files:
            print(file.as_posix())
            if(excel_name in file.as_posix()):
                weekly_scores = file.as_posix()
                return weekly_scores
                
        raise OSError("No Leaderboard/New weekly data found in file {}".format(str(current_path)))
    

        
    @staticmethod
    def LocateMoveAndClick(path,confidence=0.9):
        time.sleep(1)
        position_of_signin = pyautogui.locateCenterOnScreen(path,confidence=confidence)
        pyautogui.moveTo(position_of_signin[0],position_of_signin[1],duration=2)
        pyautogui.click(button="left")
        time.sleep(0.2)

    def add_new_page(self,link):
        pyautogui.keyDown("ctrl")
        pyautogui.press("t")
        pyautogui.keyUp("ctrl")

        #wait for the page to load
        time.sleep(1.5)

        #enter the link in the browser
        pyautogui.typewrite(link,interval=0.03)
        pyautogui.press("enter")

        #wait for the page to load
        time.sleep(2)
        
        #switch to the responses tab
        self.LocateMoveAndClick("C:/Users/yw347/Downloads/MathCommiteeAPI/AnswerDetection/objects/responses.png")