from selenium import webdriver
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

pageLoadTime = 10
numDrivers = 2
driverPath = Path(Path.cwd(), "chromedriver.exe")
service = webdriver.ChromeService(executable_path = driverPath)
drivers = [webdriver.Chrome(service=service, options = webdriver.ChromeOptions()) for i in range(numDrivers)]

#Use Threads for the low-cpu selenium drivers -- each will open up a higher-cpu chrome page process
def startPageThread(author, driverNum):
    drivers[driverNum].get('https://www.google.com')
    input("Please press enter to continue")
    print(drivers[driverNum].page_source)
    drivers[driverNum].get('https://www.lens.org/lens/search/patent/list?q=joseph%20stern')

startPageThread("test", 0)