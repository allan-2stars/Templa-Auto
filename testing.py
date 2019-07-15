
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


# # open analysis details dialouge window
contractDetailWindow = app.window(title_re='Contract - *')
contractDetailWindow.wait('exists', timeout=15)

# press the tab of QA
contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
contractDetailWindow.child_window(title="Title", auto_id="16", control_type="ComboBox").click_input()
pyautogui.typewrite('Monthly')

qaExternalItem = contractDetailWindow.window(title='2 -- External QA -- QA-EXT')
qaExternalItemExt = contractDetailWindow.window(title='4 -- QA-Ext -- QA-EXT')
qaContractItem = contractDetailWindow.window(title='2 -- Contract Cleaning -- Contract Cleaning')
existsExternalItem = qaExternalItem.exists()
existsExtItem = qaExternalItemExt.exists()
existContractItem = qaContractItem.exists()
if  existsExternalItem or existsExtItem or existContractItem:
    print("exist")

##contractDetailWindow.print_control_identifiers()


