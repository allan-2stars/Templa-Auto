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

## default filter
filterWindow.child_window(title="Default criteria").click_input()
## change the QA filters criteria
QAFilterCriteria = filterWindow.child_window(title='QA filtering criteria', control_type='TabItem')
QAFilterCriteria.click_input()

###############################################
############                    ###############
############  Basic Filtering   ###############

## filter on date range of audited date
filterWindow.child_window(auto_id="datAuditDateFrom", control_type="Edit").click_input()
pyautogui.typewrite('01062019')
filterWindow.child_window(auto_id="datAuditDateTo", control_type="Edit").click_input()
pyautogui.typewrite('30062019')

## ## click on Failed Items button to YES
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('right')
pyautogui.press('space')

## if the site is PMC and DAWR then use Ignore.
# just Default the filter will do the trick

###########   End of Basic Filtering    #########
#################################################


## change the site filters criteria
siteFilterCriteria = filterWindow.child_window(title='Site filtering criteria', control_type='TabItem')
siteFilterCriteria.click_input()
filterWindow.child_window(title="Contracts", auto_id="5", control_type="DataItem").click_input()
pyautogui.moveRel(100,0)
pyautogui.dragRel(-300,0)
pyautogui.typewrite('Affinity')
pyautogui.press('tab')

# ## Save the filter
filterWindow.Save.click_input()
