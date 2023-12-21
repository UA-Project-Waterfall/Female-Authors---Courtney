import pyautogui
import time
import webbrowser
import win32api

SCREEN_WIDTH, SCREEN_HEIGHT = pyautogui.size()
url = "www.google.com"
webbrowser.open(url)

msvcrt.getch()
position = pyautogui.position()
pyautogui.typewrite(position.x + ", " + position.y)
pyautogui.press("enter")