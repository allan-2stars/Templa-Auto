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
sheetLoader = 'QA Template' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("Reading Excel...")
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    qaTemplate = df['TEMPLATE']
    nextQaDate =  df['NEXT QA']
    status = df['STATUS']

    if status[i] == "Done" or status[i] == "Skip":
        print(siteCode[i] + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code Edit Box
    mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
    pyautogui.typewrite(siteCode[i])
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    print("contiune...")

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=15)


    # Go to QA tab
    contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
    # see if exist
    qaExternalItem = contractDetailWindow.window(title='2 -- External QA -- QA-EXT')
    qaExternalItemOther = contractDetailWindow.window(title_re='4 --*')
    print ("qa ext " + str(qaExternalItem.exists()))
    print ("qa ext other " + str(qaExternalItemOther.exists()))


    # if item exist, then see if need to change freq
    if  qaExternalItem.exists() or qaExternalItemOther.exists():

        contractDetailWindow.window(title='New version').click_input()

        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        pyautogui.PAUSE = 3.5

        # press the tab of QA
        time.sleep(5)
        contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
        if qaExternalItem.exists(): 
            qaExternalItem.click_input(double="true")
        else:
            qaExternalItemOther.click_input(double="true")
        print ("openning the qa item...")

        contractDetailWindow.child_window(title='Edit this effective version').click_input()
        pyautogui.PAUSE = 2.5

        contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
        contractQAWindow.wait('exists', timeout=15)

        ######################
        #
        # Change QA Template
        #
        ######################
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        #pyautogui.typewrite(qaTemplate[i])
        pyautogui.typewrite(qaTemplate[i])
        pyautogui.PAUSE = 2.5

        #contractQAWindow.child_window(title="Any time", control_type="RadioButton").click_input()
        contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box

        ####################################
        #
        # Quick Way to change Next QA Date
        #
        # ##################################
        # nextDateString = "13082018"
        # pyautogui.typewrite(nextDateString)
        # pyautogui.press('tab')

        #######################################
        #
        # Felxible Way to Change Next QA Date
        #
        #######################################
        # nextDateString = str(nextQaDate[i])
        # # get the date character one by one and type in
        # for letter in nextDateString:
        #     pyautogui.typewrite(letter)
        pyautogui.typewrite("01032019")
        pyautogui.press('tab')

        # Save
        contractQAWindow.Accept.click_input()
        contractDetailWindow.window(title='Request approval').click_input()
        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        print(siteCode[i] + " updated now")
        time.sleep(50)
    else:
        contractDetailWindow.Close.click_input(double=True)
        print ("Due to no external QA, closed directly.")

    print(siteCode[i] + " Done now")    
    print("########################")
    print("")
