print("Importing Utility modules")

from pathlib import Path
from random import random
from math import ceil
from turtle import position
from pandas import read_excel
from threading import Thread
from time import sleep
from datetime import datetime

print("Importing Automation modules")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webbrowser import Chrome
import pyautogui, pygetwindow, openpyxl

randomBuffer = 2
waitBuffer = 10

def __main__():
    authorTable = getAuthorTable()
    speed, numDrivers = promptDriverParams()
    threads = [Thread()] * numDrivers
    threads[0].start()
    
    outputPath = Path.joinpath(Path(__file__).parent.resolve(), "Lens Downloads " +
                               datetime.now().strftime("%m-%d %H-%M-%S"))
    drivers = initDrivers(numDrivers, outputPath)    

    for i in range(0, authorTable.shape[0]):
        index = i % numDrivers
        threads[index].join()
        args = (drivers[index], authorTable['First Name'][i], authorTable['Last Name'][i])
        threads[index] = Thread(target = pageThread, args = args)
        threads[index].start()
        sleep(speed + random() * randomBuffer)
        
    for thread in threads:
        thread.join()
        
def getAuthorTable():
    confirm = False
    while not confirm:
        authorPath = Path(input("\nName of authors excel file: "))
        if not authorPath.is_absolute():
            authorPath = Path.joinpath(Path(__file__).parent.resolve(), authorPath)
        print("File selected: " + str(authorPath))
        if(authorPath.is_file()):
            if(input("Type 'y' to confirm (Press enter otherwise): ").lower() == "y"):
                try:
                    authorTable = read_excel(authorPath)
                    confirm = True
                except:
                    print("Could not open file. Please try again.")
        else:
            print("Could not find file. Please try again.")
    return authorTable

def promptDriverParams():
    confirm = False    

    while not confirm:
        print("\nPlease enter speed below. Faster speed may be detected as unusual bot behavior.")
        print("Additionally, a random amount of time between 0-2 seconds will be added to the input speed.")
        speed = max(0.5, float(input("Base seconds per query (Recommended is 4): ").strip() or 4))
        numDrivers = ceil(waitBuffer / speed)
        print("\nSpeed set to " + str(speed) + " + [0-" + str(randomBuffer) + "] seconds per query.")
        print("Number of drivers that will be used: " + str(numDrivers))
        confirm = input("Type 'y' to confirm (Press enter otherwise): ").lower() == "y"
    return [speed, numDrivers]

def initDrivers(numDrivers, outputFolder):
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    #options.add_argument("--user-agent=Chocolate")
    options.add_experimental_option("excludeSwitches",["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")
    #options.add_argument('--blink-settings=imagesEnabled=false')
    prefs = {"profile.default_content_settings.popups": 0,
             "download.prompt_for_download": False,
             "download.default_directory": str(outputFolder),
             "download.directory_upgrade": True}
    options.add_experimental_option("prefs",prefs)
    
    drivers = [None] * numDrivers
    (SCREEN_WIDTH, SCREEN_HEIGHT) = pyautogui.size()
    
    print("\nA new Google Chrome window will open after this message for five seconds.")
    print("Press move your mouse to the position of the new tab button when it appears.")
    input("Press enter to continue.")
    print("\nOpening Google Chrome window...")
    cmdWindow = pygetwindow.getActiveWindow()
    driver = webdriver.Chrome(service = Service(), options = options)
    driver.get("https://www.google.com")
    driverWindow = pygetwindow.getActiveWindow()
    
    confirm = False
    while not confirm:
        driverWindow.activate()
        sleep(5)
        newTabPos = pyautogui.position()
        cmdWindow.activate()
        print("\nThe current mouse position has been saved. Do you want to use the current position?")
        confirm = input("Type 'y' to confirm (Press enter otherwise): ").lower() == "y"
    
    print("The program will now initialize drivers. This may take a minute.")
    input("Please do not move your mouse during this time. Press enter when you are ready.")
    driver.close()
    
    for i in range(0, numDrivers):
        driver = webdriver.Chrome(service = Service(), options = options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => false})")
        driver.get("https://www.google.com")
        
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        drivers[i] = driver
        
        pyautogui.moveTo(newTabPos[0], newTabPos[1])
        pyautogui.click()
        pyautogui.write("https://www.lens.org/lens/search/patent/list?q=" + str(int(random() * 1000)))
        pyautogui.press('enter')

    sleep(waitBuffer)
    cmdWindow.activate()
    print("\nIt is now safe to move the mouse and press keys.")
    print("Initialization is currently paused to bypass bot detection.")
    input("Please press enter once cloudfare bot-detection is passed.")
    
    print("\nAutomation will now start. Please keep driver windows open and maximized.")
    print("The drivers will automatically close once all authors are processed.")
    print("Sit back and enjoy :)")
    
    for driver in drivers:
        tabs = driver.window_handles
        driver.close()
        driver.switch_to.window(tabs[1])
        driver.outputFolder = outputFolder
    
    return drivers

def pageThread(driver, firstName, lastName):
    success = False
    while(not success):
        try:
            driver.get("https://www.lens.org/lens/search/patent/list?q=inventor.name:(" + firstName + " " + lastName + ")")
            wait = WebDriverWait(driver, waitBuffer); 
    
            button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                "body > div.lf-ui-view.ng-scope > main > section > div > div:nth-child(2) > header > " +
                "div:nth-child(3) > div.toolbar-results.clearfix.ng-scope > div > label > ul > li > a:nth-child(5)")))
            button.click()
    
            fileInput = wait.until(EC.element_to_be_clickable((By.ID, "exportFilename")))
            fileInput.send_keys(firstName + "_" + lastName)
    
            button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                "body > div.lf-ui-view.ng-scope > main > div:nth-child(4) > div.ng-scope.ng-isolate-scope > div > div > div " +
                "> div.css-modal-wrapper-inner > div > div:nth-child(1) > div.row > div.col-md-8 > form > div > button")))
            button.click()
    
            filePath = Path.joinpath(driver.outputFolder, firstName + "_" + lastName + ".csv")
            for i in range(0, waitBuffer):
                if not filePath.is_file(): sleep(1)
                else: success = True
        except:
            success = False

__main__()