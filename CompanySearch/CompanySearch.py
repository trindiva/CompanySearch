import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import os
from tkinter import *
from tkinter import filedialog
from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
import pywinauto
import random
import time
import decimal
from win32ctypes.pywin32 import win32api
import pyautogui

# open chrome and go to zoominfo.com and clickn on the login button
### NOT A USED FUNCTION! I AM GOING TO OPEN UP ZOOMINFO MANUALLY AND SET UP ALL THE PARAMS BEFORE RUNNING THE PROGRAM
def open_chrome():
    chrome_dir = r'"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"'

    chrome = Application(backend='uia')
    chrome.start(chrome_dir + ' --force-renderer-accessibility --start-maximized')

    time.sleep(float(decimal.Decimal(random.randrange(50, 150))/100))

    send_keys("zoominfo.com {ENTER 2}")

    time.sleep(float(decimal.Decimal(random.randrange(150, 300))/100))
    time.sleep(7)

    position = pyautogui.position()
    print(position)

    pywinauto.mouse.click(button='left', coords=(1737, 104))

    time.sleep(float(decimal.Decimal(random.randrange(150, 300))/100))

def send_companies():
