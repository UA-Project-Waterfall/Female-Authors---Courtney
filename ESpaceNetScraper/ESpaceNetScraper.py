print("Importing Utility modules")

from pathlib import Path
from random import random
from math import ceil
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

waitBuffer = 10

def __main__():
    authorTable = getAuthorTable()
    speed, numDrivers = promptDriverParams()
    threads = [Thread()] * numDrivers
    threads[0].start()
    
    outputPath = Path.joinpath(Path(__file__).parent.resolve(), "ESpaceNet Downloads " +
                               datetime.now().strftime("%m-%d %H-%M-%S"))
    drivers = initDrivers(numDrivers, outputPath)    

    for i in range(0, authorTable.shape[0]):
        index = i % numDrivers
        threads[index].join()
        args = (drivers[index], authorTable['First Name'][i], authorTable['Last Name'][i])
        threads[index] = Thread(target = pageThread, args = args)
        threads[index].start()
        sleep(speed)
        
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
        print("\nPlease enter speed below. Faster speed may cause lag.")
        speed = max(0.5, float(input("Base seconds per query (Recommended is 4): ").strip() or 4))
        numDrivers = ceil(waitBuffer / speed)
        print("\nSpeed set to " + str(speed) + " seconds per query.")
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
    options.add_argument('--blink-settings=imagesEnabled=false')
    prefs = {"profile.default_content_settings.popups": 0,
             "download.prompt_for_download": False,
             "download.default_directory": str(outputFolder),
             "download.directory_upgrade": True}
    options.add_experimental_option("prefs",prefs)
    
    drivers = [None] * numDrivers
    
    for i in range(0, numDrivers):
        driver = webdriver.Chrome(service = Service(), options = options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => false})")
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