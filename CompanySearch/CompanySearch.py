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
    # path to chrome
    chrome_dir = r'"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"'

    # open chrome maximized and allow access to pywinauto
    chrome = Application(backend='uia')
    chrome.start(chrome_dir + ' --force-renderer-accessibility --start-maximized')

    # sleep for random time to allow page to load and go to zoominfo.com
    time.sleep(float(decimal.Decimal(random.randrange(50, 150))/100))

    send_keys("zoominfo.com {ENTER 2}")

    # sleep for enough time for me to move mouse to the 'login' button to get coordinates
    time.sleep(float(decimal.Decimal(random.randrange(150, 300))/100))
    time.sleep(7)

    # print coordinates of mouse over the login button
    position = pyautogui.position()
    print(position)

    # move mouse to and click on the login button
    pywinauto.mouse.click(button='left', coords=(1737, 104))

    time.sleep(float(decimal.Decimal(random.randrange(150, 300))/100))

# get df with company names
def get_df():
    path = "C:\\Users\\Ivan Trindev\\ACT Capital Advisors\\DnS - Documents\\General\\DnS - Buyer Tracking List - 12.15.2019 - Sheppard Scrub.xlsx"

    df = pd.read_excel(path, sheet_name="Strategics - Final Marketing", skiprows=2)
    return df

# search for the companies
def company_search():
    # sleep for 15 sec to give me time to pull up zoominfo
    time.sleep(15)

    # get df
    df = get_df()

    # request info for the first 10 companies in the list
    count = 0
    while count < 10:
        company = df["Company Name"].iloc[count]
        send_keys(company + " {ENTER}")

        time.sleep(float(decimal.Decimal(random.randrange(200, 400))/100))

        count += 1

company_search()
