
from subprocess import Popen
from pywinauto import Desktop
from pywinauto import Application
import pyautogui
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from pywinauto.application import Application
import time
import csv
import os
import sys
import pywinauto
from datetime import datetime

## get Templa ready
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')




#templa.child_window(title='Change filter', control_type='Button').click_input()
filterWindow = templa.window(title_re='QA Completed Item Filter Detail - *')
filterWindow.wait('exists', timeout=15)

## change the site filters criteria
# siteFilterCriteria = filterWindow.child_window(title='Site filtering criteria', control_type='TabItem')
# siteFilterCriteria.click_input()

filterWindow.child_window(auto_id="cslQATemplate", control_type="Pane").click_input()
pyautogui.typewrite("276")
pyautogui.press('tab')

# filterWindow.child_window(auto_id="cslSite", control_type="Pane").click_input()

# filterWindow.child_window(auto_id="cslClient", control_type="Pane").click_input()

print('finish')


##filterWindow.print_control_identifiers()


