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
sheetLoader = 'Freq-Change' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("Reading Excel...")
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    freqNum = df['NUMBER']
    frequency = df['FREQUENCY']
    freqText = df['FREQ-TEXT']
    dayNumber = df['DAYS TO COMPLETE']
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
    pyautogui.typewrite(str(siteCode[i]))
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    print("contiune...")
    print("site code is: " + str(siteCode[i]))

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=35)


    # Go to QA tab
    contractTabWindow = contractDetailWindow.child_window(title='QA', control_type='TabItem')
    contractTabWindow.click_input()

    # see if exist, use the Task number to check, in Templa, open "Task number" tab if not there.
    qaExternalItem = contractDetailWindow.window(title='2')
    qaExternalItemOther = contractDetailWindow.window(title='4')
    print('QA other exists? ',qaExternalItemOther.exists())

    # if item exist, then see if need to change freq
    if  qaExternalItem.exists() or qaExternalItemOther.exists():
        # exist, then test if need to change
        freqEditBox = contractDetailWindow.child_window(title="Frequency", auto_id="20", control_type="ComboBox")
        freqEditBox.click_input()
        pyautogui.typewrite(freqText[i])
        pyautogui.press('tab')
        pyautogui.PAUSE = 2.5
        print("Same Frequency: " + str(qaExternalItem.exists()))

        # If standartd QA External Items not showing then, need to change
        if not (qaExternalItem.exists() or qaExternalItemOther.exists()):
            # Clear the test Text
            # freqEditBox.click_input()
            # pyautogui.press('backspace')
            # # wait again to exist
            # qaExternalItem.wait('exists', timeout=10)
            contractDetailWindow.window(title='New version').click_input(double=True)

            pyautogui.PAUSE = 2.5
            pyautogui.typewrite('y') ## equivilent to clicking "yes"
            pyautogui.PAUSE = 3.5

            # press the tab of QA
            contractTabWindow.wait('exists', timeout=35)
            contractTabWindow.click_input()
            if qaExternalItem.exists():
                qaExternalItem.click_input(double=True)
            if qaExternalItemOther.exists():
                qaExternalItemOther.click_input(double=True)
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
            # contractQAWindow.Edit10.click_input()
            # pyautogui.PAUSE = 2.5
            # pyautogui.moveRel(80, 0)
            # pyautogui.dragRel(-500,0)
            # pyautogui.typewrite(qaTemplate[i])
            # pyautogui.PAUSE = 2.5

            # Change the Freqency number
            contractQAWindow.child_window(auto_id="numFrequencyCount", control_type="Edit").click_input()
            pyautogui.typewrite(str(int(freqNum[i])))
            pyautogui.PAUSE = 2.5

            # Change the dropdown list 
            contractQAWindow.child_window(auto_id="cboFrequencyPeriod", control_type="ComboBox").click_input()
            pyautogui.typewrite(frequency[i])
            pyautogui.press("tab")
            pyautogui.PAUSE = 2.5

            contractQAWindow.child_window(auto_id="numDaysToComplete", control_type="Edit").click_input()
            pyautogui.press("delete")
            pyautogui.typewrite(str(int(dayNumber[i])))
            pyautogui.press("tab")
            pyautogui.PAUSE = 2.5

            contractQAWindow.child_window(title="Any time", control_type="RadioButton").click_input()
            contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box

            # if frequency[i] == "m":
            #     nextDateString = "01122019"
            # elif frequency[i] == "w":
            #     nextDateString = "28102019"
            # elif frequency[i] == "y":
            #     nextDateString = "01012020"
            # else:
            #     nextDateString = "01012020"
            # print('frequency text is', frequency[i])
            # print('next qa date is', nextDateString)
            # pyautogui.typewrite(nextDateString)
            # pyautogui.press('tab')
            nextDateString = str(nextQaDate[i])
        
            # get the date character one by one and type in
            for letter in nextDateString:
                pyautogui.typewrite(letter)
            
            pyautogui.press('tab')
            print('next qa date is', nextDateString)
            # Save
            contractQAWindow.Accept.click_input()
            #contractDetailWindow.print_control_identifiers()
            contractDetailWindow.child_window(title="Request approval", auto_id="[Group : workflow Tools] Tool : requestapproval - Index : 0 ", control_type="Button").click_input()
            time.sleep(2.5)
            pyautogui.typewrite('y') ## equivilent to clicking "yes"
            print(siteCode[i] + " updated now")
            time.sleep(6)
        else:
            contractDetailWindow.Close.click_input()
            print ("Due to same frequency, closed directly.")
        
    # if no qa, close it
    else:
        contractDetailWindow.Close.click_input()
        print ("No QA for this site, closed directly.")
       

    print(siteCode[i] + "Done now")
    print("##################")
    print("                  ")
    
    
