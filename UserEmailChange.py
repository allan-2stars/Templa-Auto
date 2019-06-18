from subprocess import Popen
from pywinauto import Desktop
from pywinauto import Application
import pyautogui
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from pywinauto.application import Application
import csv
import os
import sys
import pywinauto

## Start the App
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')

## start 
mainLiteUsersTab = templa.child_window(title='LITE Users', control_type='TabItem')
mainLiteUsersTab.click_input()
mainLiteUsersWindow = templa.child_window(title='LITE Users', control_type='Window')

########################
#
# Setup Excel Sheet
#
########################
excel_sheet = 'Change User Email' 
df = pd.read_excel('test.xlsx', sheet_name=excel_sheet)
print("starting...")

for i in df.index:
    userCode = df['USER CODE']
    userName = df['USER NAME']
    userEmail = df['USER EMAIL']
    status = df['STATUS']

    if status[i] == "Done":
        print(str(userName[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code Edit Box
    mainLiteUsersWindow.window(title='Name', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(userName))

    