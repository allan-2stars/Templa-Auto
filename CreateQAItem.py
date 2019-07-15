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
mainContractsTab = templa.child_window(title='QA Items', control_type='TabItem')
mainContractsTab.click_input()
mainContractsWindow = templa.child_window(title='QA Items', control_type='Window')

templa.child_window(title="New", control_type="Button").click_input()

QAItemDetailWindow = app.window(title_re='QA Item *')
QAItemDetailWindow.wait('exists', timeout=15)
print("QA Item Window opened...")

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Add QA Items' 
df = pd.read_excel('test.xlsx', sheetname=sheetLoader)
# print("Reading Excel...")
for i in df.index:
    details = df['DETAILS']
    itemGroup = df['ITEM GROUP']
    scoreCard = df['SCORE CARD']
    status = df['STATUS']

    if status[i] == "Stop":
        print("Stop here")
        break

    if status[i] == "Done" or status[i] == "Skip":
        print(details[i]+ " is Done")
        continue


    # #########################
    # # add new QA Item
    # #########################

    pyautogui.PAUSE = 2.5
    pyautogui.press('tab')
    # print('click on details')
    #QAItemDetailWindow.child_window(title="Details", control_type="Text").click_input()
    pyautogui.typewrite(details[i])
    pyautogui.press('tab')
    #QAItemDetailWindow.child_window(title="Item group", control_type="Text")
    pyautogui.typewrite(itemGroup[i])
    # pyautogui.press('tab')
    # QAItemDetailWindow.Save.click_input()
    pyautogui.PAUSE = 2.5
    print(details[i] +" Done.")
    QAItemDetailWindow.child_window(title="Save and new", control_type="Button").click()

QAItemDetailWindow.Close.click_input()
print("All QA Item Created now")
print("##################")
    
    
