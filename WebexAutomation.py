from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
import time
import ctypes
import sys


def starting_webex_automation ():
    args = sys.argv
    print(args)
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_experimental_option("detach", True)
    #chromeOptions.add_experimental_option("applicationCacheEnabled", False)
    #driver = webdriver.Chrome(chrome_options=chromeOptions)
    capabilities = {
        "browserName": "chrome",
        "version": "77",
        "enableVNC": True,
        "enableVideo": False
    }
    driver = webdriver.Remote(
    command_executor="http://10.0.0.109:4444/wd/hub", desired_capabilities=capabilities)
    wait = WebDriverWait(driver, 60)
    driver.implicitly_wait(5)
    driver.set_page_load_timeout(60)
    driver.maximize_window()
    login(driver, args[1], args[2])
    DLP(driver, args[3], args[4], args[5])
    driver.quit()


def login(driver: webdriver, username: str, password: str,):
    wait = WebDriverWait(driver, 60)
    driver.get("https://teams.webex.com/signin")
    driver.find_element_by_css_selector("input[placeholder='Email address']").send_keys(username) #types username
    driver.find_element_by_css_selector("button[alt='Next']").click()
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[ alt='Password ']"))) #waits password for element
    driver.find_element_by_css_selector("input[ alt='Password ']").send_keys(password) #types password
    driver.find_element_by_css_selector("button[id='Button1']").click()
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[alt='?']"))) #waits for "?" element
    try:
        driver.switch_to_alert().accept()
    except Exception as e:
        print(e)

def files(driver: webdriver):
    spacetextbox = driver.find_element_by_css_selector("div[role='textbox']")
    textfile = os.path.abspath("Violation.txt")  # Defines text file
    wordfile = os.path.abspath("Violation2.docx")  # Defines word file
    excelfile = os.path.abspath("Violation3.xlsx")  # Defines excel file
    file_upload = driver.find_element_by_id("messageFileUpload")
    file_upload.send_keys(textfile)  # Opens violation file
    spacetextbox.send_keys('Violation File\n')  # Sends text violation
    time.sleep(2)
    file_upload.send_keys(wordfile)  # opens violation file
    spacetextbox.send_keys('Violation File\n')  # Sends word violation
    time.sleep(2)
    file_upload.send_keys(excelfile)  # opens violation file
    spacetextbox.send_keys('Violation File\n')  # Sends word violation

def find_user(driver: webdriver, user :str, wait):
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "button[alt='Search']")))
    driver.find_element_by_css_selector("button[alt='Search']").click()  # clicks on search bar
    driver.find_element_by_css_selector("textarea[placeholder='Search']").send_keys(user)  # searches user to send DLP violation to
    driver.find_element_by_css_selector(".md-list.md-list--vertical > div.contact").click()  # clicks on the user
    time.sleep(2)

def SendTextMessage(driver: webdriver, text :str,):
    textbox = driver.find_element_by_css_selector("div[role='textbox']")
    textbox.click()
    textbox.send_keys('this is not a violation\n')  # Sends non violation
    textbox.send_keys(text+'\n')  # Sends Violation1
    time.sleep(3)

def Space(driver: webdriver, user1 :str, user2 :str, wait, text:str): #Looks for the space to perform DLP sanity in
    driver.find_element_by_css_selector("button[alt='Search']").click()  # clicks on search bar
    driver.find_element_by_css_selector("textarea[placeholder='Search']").send_keys("DLP Test")  # searches for space
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[alt='Spaces']")))  # Waits for element
    try:
        space = driver.find_element_by_css_selector("div[title='DLP Test']")  # Trying to find the Space
        space.click()
    except Exception as e:  # If it couldn't find the space it will create a new one
        driver.find_element_by_css_selector("button[alt='Cancel']").click()
        driver.find_element_by_id("global-activity-menu-button").click()
        driver.find_element_by_css_selector("div[aria-label='Create a space']").click()  # Creates a space
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[alt='Create']")))  # Waits for element
        driver.find_element_by_css_selector("input[placeholder='Space name']").send_keys("DLP Test")  # Name's the space
        driver.find_element_by_css_selector("input[aria-label='Search people by name or email']").send_keys(
            user1)  # finds user 1
        driver.find_element_by_css_selector("div[class='person-list-item-name']").click()  # adds the user
        driver.find_element_by_css_selector("input[aria-label='Search people by name or email']").send_keys(
            user2)  # finds user 2
        driver.find_element_by_css_selector("div[class='person-list-item-name']").click()  # adds the user
        driver.find_element_by_css_selector("button[alt='Create']").click()  # Creates the space
        driver.find_element_by_css_selector("button[alt='Message']").click()  # Starts chat
    finally:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Add attachment']")))  # waits for attachment element
        SendTextMessage(driver,text)
        time.sleep(3)
        files(driver)
        time.sleep(3)

def Teams(driver: webdriver, user1 :str, user2 :str, wait, text:str): #Looks for the team to perform DLP sanity in
    driver.find_element_by_css_selector("button[aria-label^='Teams']").click() #Clicks on the Teams button on the side menu
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[title='Create Team']"))) #Waits for the Create Team button
    try:
        driver.find_element_by_css_selector("div[title='DLP Team']").click() #Clicks on the DLP Team if it exists
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[aria-label='General']")))
        driver.find_element_by_css_selector("div[title='General']").click()  # Clicks on the general channel in the team
    except Exception as e:
        driver.find_element_by_css_selector("button[title='Create Team']").click()
        driver.find_element_by_css_selector("input[placeholder='Name this Team']").send_keys("DLP Team") #Names the team
        driver.find_element_by_css_selector("button[type='submit']").click() # Creates the team
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div[aria-label='General']")))
        driver.find_element_by_css_selector("button[aria-label='Members']").click() #clicks on the members tab
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Archive Team']"))) #waits for element
        driver.find_element_by_css_selector("div[title='Add Team Member'] >div >button").click()
        driver.find_element_by_css_selector("input[title='Add people by name or email']").send_keys(user1) #Searches the user
        driver.find_element_by_css_selector("div[class='person-list-item-name']").click()  # adds the user
        driver.find_element_by_css_selector("input[title='Add people by name or email']").send_keys(user2) #Searches second user
        driver.find_element_by_css_selector("div[class='person-list-item-name']").click()  # adds second user
        driver.find_element_by_css_selector("button[aria-label='Spaces']").click()
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div[aria-label='General']")))
        driver.find_element_by_css_selector("div[aria-label='General']").click()
    finally:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Add attachment']")))  # waits for attachment element
        SendTextMessage(driver,text)  # Starts DLP Sanity
        time.sleep(3)
        files(driver)
        time.sleep(3)

def DLP(driver: webdriver, user1: str, user2: str, text: str):
     wait = WebDriverWait(driver, 35)
     find_user(driver, user1, wait)
     wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Add attachment']")))  # waits for attachment element
     SendTextMessage(driver,text)
     files(driver)
     time.sleep(3)
     Space(driver, user1, user2, wait, text)
     Teams(driver, user1, user2, wait, text)


starting_webex_automation()
time.sleep(5)
print('Automation has finished succesfully.')
'''ctypes.windll.user32.MessageBoxW(0, "Automation has finished succesfully.", "Automation Test", 1)'''
