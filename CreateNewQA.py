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
# start 
mainContractsTab = templa.child_window(title='Contracts', control_type='TabItem')
mainContractsTab.click_input()
mainContractsWindow = templa.child_window(title='Contracts', control_type='Window')



########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Create QAs' 
df = pd.read_excel('test.xlsx', sheetname=sheetLoader)
# print("Reading Excel...")
for i in df.index:
    siteCode=df['CODE']
    siteName = df['SITE NAME']
    dateStart = df['DATE START']
    dateFinish = df['DATE FINISH']
    qaTemplate = df['QA TEMPLATE']
    task = df['TASK']
    freqNum = df['FREQ NUM']
    frequency = df['FREQUENCY']
    daysToComplete = df['DAYS TO COMPLETE']
    status = df['STATUS']

    if status[i] == "Stop":
        print("Stop here")
        break

    if status[i] == "Done":
        print(siteName[i]+ " is Done")
        continue


    # #########################
    # # open site window
    # #########################

    # click on the Code Edit Box
    mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(siteCode[i]))
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    print("contiune...")
    print("Site code is: " + siteCode[i])

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=15)
    ## Start a New Version

    contractDetailWindow.window(title='New version').click_input(double=True)

    # and confirm you want to start a new version
    pyautogui.PAUSE = 2.5
    pyautogui.typewrite('y') ## equivilent to clicking "yes"
    pyautogui.PAUSE = 3.5

    # press the tab of QA
    contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()

    ## click Add    
    contractDetailWindow.child_window(title="Add", auto_id="btnAddQA", control_type="Button").click_input()

    ## in Contrac QA Window
    contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
    contractQAWindow.wait('exists', timeout=15)
    contractQAWindow.child_window(auto_id="datEffectiveFrom", control_type="Edit").click_input()

    # change this as require much quicker that for loop
    dataStart = "19022019"
    pyautogui.typewrite(dataStart)
    pyautogui.press('tab')
    dateFinish = "28022019"
    pyautogui.typewrite(dateFinish)
    dateNextQA = "13022019"

    # pyautogui.press('tab')
    #nextDateString = str(nextQaDate[i])

    # get the date character one by one and type in
    # for letter in nextDateString:
    #     pyautogui.typewrite(letter)
    # pyautogui.press('tab')

    ## contractQAWindow.child_window(auto_id="datEffectiveTo", control_type="Edit")

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(qaTemplate[i])
    pyautogui.press('tab')
    ## the title will auto comes up.
    ## fiirst remove them
    ## add desired one
    pyautogui.press('backspace')
    pyautogui.typewrite(siteName[i])
    pyautogui.press('tab')
    pyautogui.typewrite(str(task[i]))

    # Change the Freqency number
    contractQAWindow.child_window(auto_id="numFrequencyCount", control_type="Edit").click_input()
    pyautogui.typewrite(str(int(freqNum[i])))
    pyautogui.PAUSE = 2.5

    # Change the dropdown list 
    contractQAWindow.child_window(auto_id="cboFrequencyPeriod", control_type="ComboBox").click_input()
    pyautogui.typewrite(frequency[i])
    pyautogui.press("tab")
    pyautogui.PAUSE = 2.5

    # Change the Days to complete boxt
    contractQAWindow.child_window(auto_id="numDaysToComplete", control_type="Edit").click_input()
    pyautogui.press("delete")
    pyautogui.typewrite(str(int(daysToComplete[i])))
    pyautogui.press("tab")
    pyautogui.PAUSE = 2.5


    contractQAWindow.child_window(title="Any time", control_type="RadioButton").click_input()
    contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box
    pyautogui.typewrite(dateNextQA)
    pyautogui.press('tab')

    contractQAWindow.Accept.click_input()
    contractDetailWindow.window(title='Request approval').click_input()
    pyautogui.PAUSE = 2.5
    pyautogui.typewrite('y') ## equivilent to clicking "yes"
    print(siteCode[i] + " updated now")
    time.sleep(16)

# contractDetailWindow.window(title='Request approval').click_input()
# pyautogui.PAUSE = 2.5
# pyautogui.typewrite('y') ## equivilent to clicking "yes"

print("All QA completed now")
print("##################")
    
    
