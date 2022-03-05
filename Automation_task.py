from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from PIL import Image
import pyautogui
import os
import win32gui
import time
from win32com import client
import ait

def load_website(driver, url):
    website_load_count = 0
    website_load_status = True
    while True:
        try:
            driver.get(url)
            print("Opened")
            break
        except TimeoutException as e:
            print("Another iteration")
            time.sleep(5)
            website_load_count += 1
            if website_load_count == 3:
                print("Failed Website opening")
                website_load_status = False
                break
    return website_load_status

def Select_open_option(driver):
    driver.maximize_window()
    #send ctr + o 
    try:
        a = ActionChains(driver)
        # perform the ctrl+c pressing action
        a.key_down(Keys.CONTROL).send_keys('O').key_up(Keys.CONTROL).perform()
    except:
        print("Control + O keys failed to send")
    
    return driver


def macro_handler():
    obj = client.Dispatch("SeeShell")
    obj.open(True, 90)
    img_path = r"C:\Users\Dell\Documents\Redbuffer_intern\3.png"
    i = obj.setVariable("path",img_path)
    i = obj.echo("upload_image")
    i = obj.play("upload_image")
    return i,obj

    
    
    
def select_cartoonizer(driver):
    time.sleep(3)
    Open_option = driver.find_element_by_xpath("//*[@id='digitalart_ART_0001_menu']") 
    Open_option.click()   
    return driver


def screenshot(driver):
    ait.move(22,15)
    time.sleep(15)
    driver.save_screenshot("C:\\Users\Dell\\Documents\\screenshots\\1.png")  #path of saved image
#     screenshot = Image.open("1.png")
#     screenshot.show()
    return driver
    
url = 'https://www.befunky.com/create/photo-to-cartoon/'
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(chrome_options=options, executable_path=r'C:\Users\Dell\Documents\dataScienceProjects\other\chromedriver.exe')
img_path = "C:\\Users\\Dell\\Documents\\Redbuffer_intern\\2.png"
website_load_status = load_website(driver, url)
print("website_load_status :",website_load_status)

driver = Select_open_option(driver)
i,obj = macro_handler()
driver = select_cartoonizer(driver)
driver = screenshot(driver)
obj.close()
driver.close()
print("Ended")