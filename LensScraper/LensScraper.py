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
from tqdm import tqdm
import pyautogui, openpyxl, csv

randomBuffer = 2
waitBuffer = 4
downloadBuffer = 8
defFile = "doctor_data 1.xlsx"

def __main__():
    [authorTable, start, end] = getAuthorTable()
    speed, numDrivers = promptDriverParams()
    threads = [Thread()] * numDrivers
    threads[0].start()
    
    outputPath = Path.joinpath(Path(__file__).parent.resolve(), "Lens Downloads " +
                               datetime.now().strftime("%m-%d %H-%M-%S"))
    drivers = initDrivers(numDrivers, outputPath)    
    
    for i in tqdm(range(int(start) - 2, int(end) - 1)):
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
        authorPath = Path(input("\nName of authors excel file (Leave blank to use " + defFile + "): ") or defFile)
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
    
    print("\nPlease enter the information below based on the excel row indices.")
    start = input("Starting row number (Use 2 for first element): ")
    end = input("Ending row number (Leave blank to set to end): ") or len(authorTable) + 1
    
    return [authorTable, start, end]

def promptDriverParams():
    confirm = False    

    while not confirm:
        print("\nPlease enter speed below. Faster speed may be detected as unusual bot behavior.")
        print("Additionally, a random amount of time between 0-2 seconds will be added to the input speed.")
        speed = max(0.5, float(input("Base seconds per query (Recommended is 4): ").strip() or 4))
        numDrivers = ceil((downloadBuffer + waitBuffer) / speed)
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
    
    print("\nA new Google Chrome window will open after this message.")
    print("Please click the new tab button to open a new tab when it appears.")
    input("After clicking, keep the mouse on the button. Press enter to continue.")
    print("\nOpening Google Chrome window...")
    
    testDriver = webdriver.Chrome(service = Service(), options = options)
    test_html = "<html><head></head><body><div>Please click the new tab button.</div></body></html>"
    testDriver.get("data:text/html;charset=utf-8," + test_html)
    testDriver.execute_script('document.title = "Press New Tab Button >>>"')
    
    confirm = False
    while not confirm:
        while len(testDriver.window_handles) == 1:
            testDriver.maximize_window()
        newTabPos = pyautogui.position()
        tabs = testDriver.window_handles
        testDriver.switch_to.window(tabs[1])
        testDriver.close()
        testDriver.switch_to.window(tabs[0])
        testDriver.minimize_window()
        print("\nThe current mouse position has been saved. Do you want to use the current position?")
        confirm = input("Type 'y' to confirm (Press enter otherwise): ").lower() == "y"
    
    print("\nThe program will now initialize drivers. Please do not move the mouse during this time.")
    print("A window will appear to let you know when it is safe to move the mouse.")
    input("Press enter when you are ready.")
    
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

    sleep(downloadBuffer)
    test_html = "<html><head></head><body><div>It is safe to move the mouse now.\n" + \
        "Please navigate back to the command window.</div></body></html>"
    testDriver.get("data:text/html;charset=utf-8," + test_html)
    testDriver.execute_script('document.title = "Navigate to CMD window."')
    testDriver.maximize_window()
    
    print("\nInitialization is currently paused to bypass bot detection.")
    input("Please press enter once cloudfare bot-detection is passed.")
    
    print("\nAutomation will now start. Please keep driver windows open and maximized.")
    print("The drivers will automatically close once all authors are processed.")
    print("Sit back and enjoy :)")
    testDriver.close()
    
    for driver in drivers:
        tabs = driver.window_handles
        driver.close()
        driver.switch_to.window(tabs[1])
        driver.outputFolder = outputFolder
    
    return drivers

def pageThread(driver, firstName, lastName):
    attempt = 0
    fileName = (firstName + " " + lastName).replace(" ", "_")
    filePath = Path.joinpath(driver.outputFolder, fileName + ".csv")
    while(attempt < 3):
        try:
            driver.get("https://www.lens.org/lens/search/patent/list?q=inventor.name:(" + firstName + " " + lastName + ")")
            wait = WebDriverWait(driver, downloadBuffer); 
    
            button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                "body > div.lf-ui-view.ng-scope > main > section > div > div:nth-child(2) > header > " +
                "div:nth-child(3) > div.toolbar-results.clearfix.ng-scope > div > label > ul > li > a:nth-child(5)")))
            
            foundNone = driver.find_elements(By.CSS_SELECTOR, "body > div.lf-ui-view.ng-scope > main > section > div > section > " +
                "main > div:nth-child(1) > section > div.error--notice--text > h2")
            if len(foundNone) > 0 and foundNone[0].text == "Your search did not match any documents":
                with open(filePath, 'w', newline = "") as file:
                    writer = csv.writer(file)
                    writer.writerow(["Search did not match any documents"])
                attempt = 5
            else:
                button.click()
    
                fileInput = wait.until(EC.element_to_be_clickable((By.ID, "exportFilename")))
                fileInput.send_keys(fileName)
                sleep(waitBuffer)
                fileInput.send_keys(Keys.ENTER)
    
                for i in range(0, downloadBuffer):
                    if not filePath.is_file(): sleep(1)
                    else:
                        attempt = 5
            
                if attempt < 5:
                    button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                            "body > div.lf-ui-view.ng-scope > main > div:nth-child(4) > div.ng-scope.ng-isolate-scope > div > div > div " +
                            "> div.css-modal-wrapper-inner > div > div:nth-child(1) > div.row > div.col-md-8 > form > div > button")))
                    button.click()
                
                for i in range(0, downloadBuffer):
                    if not filePath.is_file(): sleep(1)
                    else:
                        attempt = 5
        except:
            attempt += 1
    if attempt < 5:
        print("\nFailed at: " + firstName + " " + lastName + "\n")
__main__()