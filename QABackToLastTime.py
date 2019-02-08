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
sheetLoader = 'BackToLastQA' 
df = pd.read_excel('test.xlsx', sheetname=sheetLoader)
print("Reading Excel...")
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    status = df['STATUS']

    if status[i] == "Done":
        print(siteCode[i] + " is Done")
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
    print("site code is: " + siteCode[i])

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=15)


    # Go to QA tab
    contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
    # see if exist
    qaExternalItem = contractDetailWindow.window(title='2 -- External QA -- QA-EXT')
    qaExternalItemOther = contractDetailWindow.window(title='4 -- QA-Ext -- QA-EXT')


    # if item exist, then see if need to change freq
    if  qaExternalItem.exists() or qaExternalItemOther.exists():
        # exist, then test if need to change
        contractDetailWindow.window(title='New version').click_input(double=True)

        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        time.sleep(5)

        # press the tab of QA
        contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
        qaExternalItem.click_input(double=True)
        print ("openning the qa item...")

        contractDetailWindow.child_window(title='Edit this effective version').click_input()
        pyautogui.PAUSE = 2.5

        contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
        contractQAWindow.wait('exists', timeout=15)


        # # If needed, un-comment below function 
        # # for change QA template at the same time.
        
        # ######################
        # #
        # # Change QA Template
        # #
        # ######################

        
        contractQAWindow.child_window(auto_id="datLastQA", control_type="Edit").click_input()
        pyautogui.hotkey('ctrl','c')

        contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('tab')

        # Save
        contractQAWindow.Accept.click_input()
        contractDetailWindow.window(title='Request approval').click_input()
        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        print(siteCode[i] + " updated now")
        time.sleep(16)
        
    # if no qa, close it
    else:
        contractDetailWindow.Close.click_input()
        print ("No QA for this site, closed directly.")
       

    print(siteCode[i] + "Done now")
    print("##################")
    print("                  ")
    
    
