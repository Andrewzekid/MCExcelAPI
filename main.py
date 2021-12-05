from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
driver = webdriver.Chrome()
driver.get("https://outlook.office.com/mail/inbox")
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
    
    
     
