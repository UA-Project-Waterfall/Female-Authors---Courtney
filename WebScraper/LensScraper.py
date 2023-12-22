from pathlib import Path
import time, random, threading
import pyautogui, pandas

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

authorFile = input("Name of author list file: ")
pageLoadTime = 8
speed = max(0.5, float(input("Seconds per query (Default is 2): ").strip() or 2))
numDrivers = round(pageLoadTime / speed)
print("Speed adjusted to " + str(pageLoadTime / numDrivers) + " seconds per query")
maxRandTime = 2

def __main__():
    authorTable = pandas.read_excel(Path.joinpath(Path(__file__).parent.resolve(), authorFile))
    threads = [threading.Thread()] * numDrivers
    threads[0].start()
    
    drivers = initDrivers()    

    for i in range(0, authorTable.shape[0]):
        index = i % numDrivers
        threads[index].join()
        args = (drivers[index], authorTable['First Name'][i] + " " + authorTable['Last Name'][i])
        threads[index] = threading.Thread(target = pageThread, args = args)
        threads[index].start()
        time.sleep(pageLoadTime / numDrivers + random.random() * maxRandTime)
        
def waitForTag(driver, tag):
    while(not driver.find_elements(By.TAG_NAME, tag)):
        time.sleep(1)

def initDrivers():
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized") 
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches",["enable-automation"])
    
    drivers = [None] * numDrivers
    (SCREEN_WIDTH, SCREEN_HEIGHT) = pyautogui.size()

    for i in range(0, numDrivers):
        driver = webdriver.Chrome(service = Service(), options = options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => false})")
        driver.get("https://www.google.com")
        waitForTag(driver, "body")
        drivers[i] = driver
        
        pyautogui.moveTo(SCREEN_WIDTH * 0.17, SCREEN_HEIGHT * 0.02)
        pyautogui.click()
        pyautogui.write("https://www.lens.org/lens/search/patent/list?q=" + str(int(random.random() * 1000)))
        pyautogui.press('enter')

    time.sleep(pageLoadTime)
    for driver in drivers:
        tabs = driver.window_handles
        driver.close()
        driver.switch_to.window(tabs[1])
    
    return drivers

def pageThread(driver, query):
    driver.get("https://www.lens.org/lens/search/patent/list?q=" + query)
    time.sleep(pageLoadTime)
    waitForTag(driver, "body")
    htmlText = driver.execute_script("return document.body.innerHTML;")
    processHTML(htmlText)

def processHTML(htmlText):
    print(htmlText)

__main__()