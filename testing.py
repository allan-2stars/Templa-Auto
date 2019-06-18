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
    userGroup = df['USER GROUP']
    status = df['STATUS']

    if status[i] == "Done":
        print(str(userName[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break


    ## start 
    mainUsersTab = templa.child_window(title=userGroup[i], control_type='TabItem')
    mainUsersTab.click_input()
    mainUsersWindow = templa.child_window(title=userGroup[i], control_type='Window')

    # click on the Code Edit Box
    mainUsersWindow.window(title='Name', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(userName[i]))

    userEmailExists = mainUsersWindow.child_window(title=str(userEmail[i]), control_type="DataItem")
    
    if userEmailExists.exists():  
        print("User Name: " + userName[i])
        print("Already assigned to " + userEmail[i])
        print("#################################")
        print(" ")
        pyautogui.moveRel(-25, 25) 
        pyautogui.click() # reset the select status

    else:
        print("Email is incorrect, ready to change")
        pyautogui.moveRel(-25, 25) 
        pyautogui.doubleClick() # open the users window by double click

        userDetailWindow = app.window(title_re='User Details - *')
        userDetailWindow.wait('exists', timeout=15)
        # userDetailWindow.window(title='General', control_type='TabItem').click_input()
        # print("found the users General Tab: " + str(userName[i]))
        userDetailWindow.print_control_identifiers()


