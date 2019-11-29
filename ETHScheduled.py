from subprocess import Popen
from pywinauto import Desktop
from pywinauto import Application
import pyautogui
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from pywinauto.application import Application
import time
from datetime import datetime
import calendar
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
# mainContractsTab = templa.child_window(title='Contracts', control_type='TabItem')
# mainContractsTab.click_input()
# mainContractsWindow = templa.child_window(title='Contracts', control_type='Window')

# #########################
# # open ETH site window
# #########################

# # click on the Code Edit Box
# mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
# pyautogui.typewrite("VI-THO01")
# pyautogui.moveRel(0, 25) 
# pyautogui.doubleClick() # open the site by double click

print("contiune...")
print("Site Activated")
next_month = datetime.now().month + 1
current_year = datetime.now().year
lastday_next_month = calendar.monthrange(current_year, next_month)[1]

dateStartString = '01' + str(next_month) + str(current_year)
dateEndString = str(lastday_next_month) + str(next_month) + str(current_year)

# # open analysis details dialouge window
contractDetailWindow = app.window(title_re='Contract - *')
contractDetailWindow.wait('exists', timeout=15)
## Start a New Version

# contractDetailWindow.window(title='New version').click_input(double=True)

## and confirm you want to start a new version
# pyautogui.PAUSE = 2.5
# pyautogui.typewrite('y') ## equivilent to clicking "yes"
# pyautogui.PAUSE = 3.5

# press the tab of QA
contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()


########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'ETH' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
# print("Reading Excel...")
for i in df.index:
    area = df['AREA']
    title = df['TITLE']
    dateStart = df['DATE START']
    dateEND = df['DATE END']
    qaTemplate = df['QA TEMPLATE']
    task = df['TASK']
    status = df['STATUS']
    monthJan = df['Jan']
    monthFeb = df['Feb']
    monthMar = df['Mar']
    monthApr = df['Apr']
    monthMay = df['May']
    monthJun = df['Jun']
    monthJul = df['Jul']
    monthAug = df['Aug']
    monthSep = df['Sep']
    monthOct = df['Oct']
    monthNov = df['Nov']
    monthDec = df['Dec']
    completeTitle = area[i] + ' - ' + title[i]

    print(completeTitle + str(status[i]))

    use_month = monthJan   
    if next_month == 2:
        use_month = monthFeb
    if next_month == 3:
        use_month = monthMar
    if next_month == 4:
        use_month = monthApr
    if next_month == 5:
        use_month = monthMar
    if next_month == 6:
        use_month = monthJun
    if next_month == 7:
        use_month = monthMJul
    if next_month == 8:
        use_month = monthAug
    if next_month == 9:
        use_month = monthSep
    if next_month == 10:
        use_month = monthOct
    if next_month == 11:
        use_month = monthNov
    if next_month == 12:
        use_month = monthDec
        
    # 'x' marks need to set it up, otherwise no need setup.
    if use_month[i] != "x":
        print("the QA for month of " + str(next_month))
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    if status[i] == "Done" or status[i] == "Skip":
        print(completeTitle + " is Done")
        continue

    ## click Add    
    contractDetailWindow.child_window(title="Add", auto_id="btnAddQA", control_type="Button").click_input()

    ## in Contrac QA Window
    contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
    contractQAWindow.wait('exists', timeout=15)
    contractQAWindow.child_window(auto_id="datEffectiveFrom", control_type="Edit").click_input()

    pyautogui.typewrite(dateStartString)
    pyautogui.press('tab')
    pyautogui.typewrite(dateEndString)
    #nextDateString = str(nextQaDate[i])

    # get the date character one by one and type in
    # for letter in nextDateString:
    #     pyautogui.typewrite(letter)
    pyautogui.press('tab')

    ## contractQAWindow.child_window(auto_id="datEffectiveTo", control_type="Edit")
    pyautogui.press('tab')
    pyautogui.typewrite(qaTemplate[i])
    pyautogui.press('tab')
    ## the title will auto comes up.
    ## add desired one
    pyautogui.typewrite(completeTitle)
    pyautogui.press('tab')
    pyautogui.typewrite(str(task[i]))
    pyautogui.press('tab')
    contractQAWindow.child_window(title="Any time", control_type="RadioButton").click_input()
    contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box
    pyautogui.typewrite(dateStartString)
    pyautogui.press('tab')

    contractQAWindow.Accept.click_input()
    print("QA done: " + completeTitle)

# contractDetailWindow.window(title='Request approval').click_input()
# pyautogui.PAUSE = 2.5
# pyautogui.typewrite('y') ## equivilent to clicking "yes"

print("All QA completed now")
print("##################")
    
    
