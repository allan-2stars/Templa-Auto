
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
contractDetailWindow.child_window(title="Add", auto_id="btnAddQA", control_type="Button").click_input()

## in Contrac QA Window
dlg = app.top_window()
dlg.print_control_identifiers()


