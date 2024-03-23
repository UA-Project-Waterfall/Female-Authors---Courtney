print("Importing Utility modules")

from logging import root
from pathlib import Path
from random import random
from math import ceil
from tracemalloc import start
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

defFile = "doctor_data 1.xlsx"
downloadBuffer = 8
waitBuffer = 4

def __main__():
    [authorTable, start, end] = getAuthorTable()
    numDrivers = promptDriverParams()
    
    outputPath = Path.joinpath(Path(__file__).parent.resolve(), "Lens Downloads " +
                               datetime.now().strftime("%m-%d %H-%M-%S"))
    drivers = initDrivers(numDrivers, outputPath)
    manager = LensManager(drivers = drivers, outputPath = outputPath)    
    indices = range(int(start) - 2, int(end) - 1)
    pbar = tqdm(indices)
    for i in indices:
        manager.addQuery(authorTable['First Name'][i] + " " + authorTable['Last Name'][i])
    manager.start()
    while manager.running:
        pbar.n = manager.queryIndex
        pbar.update()
    manager.join()
        
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
        numDrivers = int(input("Number of drivers to use (Default is 3): "))
        confirm = input("Type 'y' to confirm (Press enter otherwise): ").lower() == "y"

    return numDrivers

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
    
    (SCREEN_WIDTH, SCREEN_HEIGHT) = pyautogui.size()
    print("\nOpening Google Chrome window...")
    
    testDriver = webdriver.Chrome(service = Service(), options = options)
    test_html = "<html><head></head><body><div>Please click the new tab button.</div></body></html>"
    testDriver.get("data:text/html;charset=utf-8," + test_html)
    testDriver.execute_script('document.title = "Press New Tab Button >>>"')
    
    while len(testDriver.window_handles) == 1:
        testDriver.maximize_window()
    newTabPos = pyautogui.position()
    tabs = testDriver.window_handles
    testDriver.switch_to.window(tabs[1])
    testDriver.close()
    testDriver.switch_to.window(tabs[0])
    testDriver.minimize_window()
    
    #return [testDriver]    #uncomment to truncate setup for debug
    
    print("\nThe program will now initialize drivers. Please do not move the mouse during this time.")
    print("A window will appear to let you know when it is safe to move the mouse.")
    input("Press enter when you are ready.")
    
    drivers = [None] * numDrivers
    for i in range(0, numDrivers):
        driver = webdriver.Chrome(service = Service(), options = options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => false})")
        driver.get("https://www.google.com")
        
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        drivers[i] = driver
        
        pyautogui.moveTo(newTabPos[0], newTabPos[1])
        pyautogui.click()
        pyautogui.write("https://www.lens.org/lens/search/patent/list?q=" + str(int(random() * 1000)))
        pyautogui.press('enter')

    sleep(waitBuffer)
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
    
    return drivers
       
class LensManager:
    drivers = []
    queries = []
    outputPath = Path(__file__).parent.resolve()
    queryIndex = 0
    downloadBuffer = 10
    running = False
    
    loadThread = Thread()
    saveThread = Thread()
    
    def __init__(self, minDelay = 1, maxDelay = 3, drivers = [], outputPath : Path = None):
        def randPause():
            sleep(minDelay + random() * (maxDelay - minDelay))
        self.randPause = randPause
        self.loadThread = Thread(target = self.mainLoadThread, args = [])
        self.saveThread = Thread(target = self.mainSaveThread, args = [])
        self.outputPath = outputPath
        outputPath.mkdir(parents = True, exist_ok = True)
        self.addDrivers(drivers)
    
    def addDriver(self, driver):        
        self.addDrivers([driver])
    
    def addDrivers(self, drivers):
        for driver in drivers:
            driver.wait = WebDriverWait(driver, 10)
            driver.query = None
            driver.filePath = None
            driver.foundNone = False
            driver.saveReady = False
            driver.loadReady = True
        self.drivers += drivers
        
    def addQueries(self, queries):
        self.queries += queries
    def addQuery(self, query):
        self.queries += [query]
    
    def mainLoadThread(self):
        while self.running:
            while(self.running and self.queryIndex < len(self.queries)):
                sleep(1)
                for driver in self.drivers:
                    if driver.loadReady:
                        driver.loadReady = False
                        driver.query = self.queries[self.queryIndex]
                        driver.filePath = Path.joinpath(self.outputPath, driver.query.replace(" ", "_") + ".csv")
                        self.queryIndex += 1
                        self.randPause()
                        Thread(target = self.__loadThread, args = [driver]).start()
                        self.randPause()
                        break
    
    def mainSaveThread(self, inProgress = True):
        while(self.running or inProgress):
            sleep(1)
            inProgress = False
            for driver in self.drivers:
                inProgress = inProgress or not driver.loadReady
                if driver.saveReady:
                    driver.saveReady = False
                    self.randPause()
                    self.__saveThread(driver)
                    self.randPause()

    def __loadThread(self, driver):
        error = ""
        for i in range (0, 3):
            try:
                driver.get("https://www.lens.org/lens/search/patent/list?q=inventor.name:(" + driver.query + ")")
                foundNone = driver.find_elements(By.CSS_SELECTOR, "body > div.lf-ui-view.ng-scope > main > section > div > section > " +
                    "main > div:nth-child(1) > section > div.error--notice--text > h2")
                if len(foundNone) > 0 and foundNone[0].text == "Your search did not match any documents":
                    with open(driver.filePath, 'w', newline = "") as file:
                        writer = csv.writer(file)
                        writer.writerow(["Search did not match any documents"])
                    driver.loadReady = True
                    return
                driver.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "body > div.lf-ui-view.ng-scope > main > section > div > div:nth-child(2) > header > " +
                    "div:nth-child(3) > div.toolbar-results.clearfix.ng-scope > div > label > ul > li > a:nth-child(5)"))).click()
                break
            except Exception as e: error = e
        if error != "":
            print("\nFailed to load:" + driver.query)
            print("\nError message:", error)
            #self.startQuery(driver.query)
            driver.loadReady = True
        else: driver.saveReady = True
    
    def __saveThread(self, driver):
        error = ""
        for i in range (0, 3):
            try:
                fileInput = driver.wait.until(EC.element_to_be_clickable((By.ID, "exportFilename")))
                fileInput.send_keys(driver.query.replace(" ", "_"))
                self.randPause()
                button = driver.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                        "body > div.lf-ui-view.ng-scope > main > div:nth-child(4) > div.ng-scope.ng-isolate-scope > div > div > div " +
                        "> div.css-modal-wrapper-inner > div > div:nth-child(1) > div.row > div.col-md-8 > form > div > button")))
                    
                while not driver.filePath.is_file():
                    button.click()
                    for i in range(0, self.downloadBuffer):
                        if not driver.filePath.is_file(): sleep(1)
                        else: break
                break
            except Exception as e: error = e
        if error != "":
            print("\nFailed to save:", driver.query)
            print("\nError message:", error)
        driver.loadReady = True
    
    def randPause(self): pass

    ''' Stops any new queries. Allows in-progress queries to finish '''
    def stop(self):
        self.running = False
    
    ''' Starts manager if not already on '''
    def start(self):
        self.running = True
        if not self.loadThread.is_alive(): self.loadThread.start()
        if not self.saveThread.is_alive(): self.saveThread.start()

    ''' Waits until the last query and download is finished '''
    def join(self):
        self.running = False
        self.loadThread.join()
        self.saveThread.join()
__main__()