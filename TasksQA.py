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


from functions.functions_utils import tm_init

## get the appliation handler from the init function
templa = tm_init()[0]
app = tm_init()[1]

print("Starting...")
## start 
mainContractsTab = templa.child_window(title='Contracts', control_type='TabItem')
mainContractsTab.click_input()
mainContractsWindow = templa.child_window(title='Contracts', control_type='Window')

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Chagne Tasks' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("Reading Excel...")
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    taskName = df['TASK']
    taskNumber = df['TASK NUMBER']
    status = df['STATUS']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(siteCode[i]) + " is Done")
        continue

    if status[i] == "Skip":
        print(str(siteCode[i]) + " is Sklipped")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code Edit Box
    mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(siteCode[i]))
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    print("contiune...")
    print("site code is: " + str(siteCode[i]))

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=15)


    # Go to QA tab
    contractDetailWindow.child_window(title='Tasks', control_type='TabItem').click_input()
    # see if exist
    # qaExternalItem = contractDetailWindow.window(title='2 -- External QA -- QA-EXT')
    
    contractDetailWindow.child_window(title='Task', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(int(taskNumber[i])))
    qaAgedCareItem = contractDetailWindow.window(title=str(taskName[i]))
    
    # qaExternalItemOther = contractDetailWindow.window(title='4 -- QA-Ext -- QA-EXT')


    # if item exist, then see if need to change freq
    if  qaAgedCareItem.exists():
        contractDetailWindow.Close.click_input()
        print ('Task already exists, closed directly.')
        print('No change')
    else:
        contractDetailWindow.window(title='New version').click_input(double=True)
        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking 'yes'
        pyautogui.PAUSE = 3.5

        # press the tab of QA
        contractDetailWindow.child_window(title='Tasks', control_type='TabItem').click_input()
        contractDetailWindow.child_window(title='Task', control_type='ComboBox').click_input()
        pyautogui.typewrite(str(int(taskNumber[i])))
        pyautogui.moveRel(0, 15)
        pyautogui.doubleClick()
        print ('openning the task item...')
        ## open the Contract Task window
        taskWindow = contractDetailWindow.window(title_re='Contract Task - *')
        taskWindow.wait('exists', timeout=15)
        ## type in task Description
        # taskWindow.child_window(auto_id="txtDescription", control_type="Edit").click_input()
        # pyautogui.moveRel(100,0)
        # pyautogui.click()
        # pyautogui.dragRel(-500,0)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(str(taskName[i]))

        ## type in task details
        # taskWindow.child_window(auto_id="txtTaskDetails", control_type="Edit").click_input()
        # pyautogui.moveRel(100,0)
        # pyautogui.click()
        # pyautogui.dragRel(-500,0)
        pyautogui.press('tab')
        pyautogui.typewrite(str(taskName[i]))
        ## Accept
        taskWindow.Accept.click_input()
        contractDetailWindow.window(title='Request approval').click_input()
        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking 'yes'

        print('site name', siteName[i] + ' Updated.')
        time.sleep(6)

    print('##################')
    print('                  ')
    
    
