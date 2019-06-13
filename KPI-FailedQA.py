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


## start 
# completedQAWindow = templa.child_window(title='Completed QA Items', control_type='TabItem')
# completedQAWindow.click_input()

# templa['Change filter'].click_input()
filterWindow = templa.window(title_re='QA Completed Item Filter Detail - *')
filterWindow.wait('exists', timeout=15)

## change the filters criteria
siteFilterCriteria = filterWindow.child_window(title='Site filtering criteria', control_type='TabItem')
siteFilterCriteria.click_input()
filterWindow.child_window(title="Contracts", auto_id="5", control_type="DataItem").click_input()
pyautogui.moveRel(100,0)
pyautogui.dragRel(-300,0)
pyautogui.typewrite('Affinity')
pyautogui.press('tab')

## Save the filter
filterWindow.Save.click_input()
