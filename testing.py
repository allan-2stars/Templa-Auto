
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

filterWindow = templa.window(title_re='QA Completed Item Filter Detail - *')

protertyFilterCriteria = filterWindow.child_window(title='Property filtering criteria', control_type='TabItem')
protertyFilterCriteria.click_input()

groupItem = filterWindow.child_window(title="Group", auto_id="2", control_type="DataItem")
matchTypeSection = groupItem.child_window(title="Match type", auto_id="3", control_type="ComboBox")
valueSection = groupItem.child_window(title="Value", auto_id="1", control_type="ComboBox")


matchTypeSection.click_input()
pyautogui.typewrite("e")
pyautogui.press("tab")
time.sleep(2)
valueSection.click_input()
pyautogui.typewrite("r")
pyautogui.press("tab")



print('finish')


#filterWindow.print_control_identifiers()


