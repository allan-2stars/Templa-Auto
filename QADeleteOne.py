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


if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')

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
sheetLoader = 'QA-Delete-Contract' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("Reading Excel...")
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    title = df['TITLE']
    status = df['STATUS']
    check_title = df['CHECK-TITLE']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(siteCode[i]) + " is Done")
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
    contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
    # get the qa title from Spread sheet
    qaTitle_to_typein = title[i]
    qaTitle_to_check = check_title[i]
    qa_item = contractDetailWindow.window(title=qaTitle_to_check)

    # type title in edit box, check if qa item exists
    title_edit_box = contractDetailWindow.child_window(title="Title", control_type="ComboBox")
    title_edit_box.click_input()
    pyautogui.typewrite(qaTitle_to_typein)
    #pyautogui.press('tab')
    pyautogui.PAUSE = 2.5
    print("QA Item Exist? " + str(qa_item.exists()))

    ## if exists
    if qa_item.exists():
        #qa_item.click_input()
        contractDetailWindow.window(title='New version').click_input(double=True)

        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        pyautogui.PAUSE = 2.5

        # press the tab of QA
        contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
        title_edit_box.click_input()
        pyautogui.typewrite(qaTitle_to_typein)
        qa_item.click_input()

        ## click Add    
        contractDetailWindow.child_window(title="Remove", auto_id="btnRemoveQA", control_type="Button").click_input()

        #contractDetailWindow.Remove.click()
        # confirm_window = contractDetailWindow.child_window(title='Confirm action')
        # print('Confirm Action Exist? ',confirm_window.exists())
        # confirm_window.Yes.click_input()
        time.sleep(1)
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        # Request Approval
        contractDetailWindow.child_window(title="Request approval", auto_id="[Group : workflow Tools] Tool : requestapproval - Index : 0 ", control_type="Button").click_input()
        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        print(str(siteCode[i]) + " updated now")
        time.sleep(6)

    else:
        contractDetailWindow.Close.click_input()
        print ("Due to QA Item NOT exists, closed directly.")


    print(str(siteCode[i]) + " Done now")
    print("##################")
    print("                  ")
    
    
print("All Done now")
print("#####################")
