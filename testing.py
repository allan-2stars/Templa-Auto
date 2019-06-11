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


def clearTextBySelectAll():
    pyautogui.moveRel(30, 0)
    pyautogui.dragRel(-500, 0, 1, button='left')
    pyautogui.press('del')

def saveAsExcel(window, pathName, folderName, fileName):
    ## export to excel and save
    window.child_window(title="Excel", auto_id="[Group : report Tools] Tool : Report_Excel - Index : 2 ", control_type="Button").click_input()
    saveAsWindow = window.child_window(title='Save As')
    saveAsWindow.wait('exists', timeout=15)
    print('save as window open')
    addressBar = saveAsWindow.child_window(title_re="Address: *", control_type="ToolBar")
    addressBar.click_input()
    pyautogui.typewrite(pathName)
    pyautogui.press('enter')
    ## add a new folder if not exists
    folderNameNeeded = saveAsWindow.child_window(title=folderName, control_type="ListItem")
    if not folderNameNeeded.exists():
        print('folder NOT exists yet.')
        saveAsWindow.child_window(title="New folder", auto_id="{E44616AD-6DF1-4B94-85A4-E465AE8A19DB}", control_type="Button").click_input()
        pyautogui.typewrite(folderName)
        pyautogui.press('enter')
    ## get into the newly created folder
    folderNameNeeded.click_input(button='left', double=True)
    ## File name type
    saveAsWindow.child_window(title="File name:", auto_id="FileNameControlHost", control_type="ComboBox").click_input()
    pyautogui.typewrite(fileName)
    ## Save button click
    saveAsWindow.child_window(title="Save", auto_id="1", control_type="Button").click_input()
    

#print(templa)
#def generate_data_file(t_interval, interface_name, file_name):
# start Wireshark
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')

## start a report with title 
analysisWindow = app.window(title='Affinity Monthly')


analysisWindow['Select live report'].click_input()

## open site selection window
reportConfigWindow = analysisWindow.child_window(title='QA Analysis Report Configurations')
reportConfigWindow.wait('exists', timeout=15)

## type report title 
reportConfigWindow.window(title='Description', control_type='ComboBox').click_input()
pyautogui.typewrite('Affinity Monthly')
pyautogui.moveRel(0, 25) 
pyautogui.click() # open the site by double click
reportConfigWindow.Select.click_input()

 analysisWindow['Run report'].click_input()
print("loading data...")
## Header defined
dragArea = analysisWindow.child_window(auto_id="GroupByBox", control_type="Group")
siteHeader = analysisWindow.child_window(title="Site", auto_id="5", control_type="ComboBox")
qaItemHeader = analysisWindow.child_window(title="QA Item",  control_type="ComboBox")

## Drag area title defined
siteDragArea = analysisWindow.child_window(title="Site", control_type="Button")
qaItemDragArea = analysisWindow.child_window(title="QA Item", control_type="Button")

## drag "Site" Label up
siteHeader.click_input(button='left', double='true')
pyautogui.PAUSE = 2.5
pyautogui.moveRel(0, -20)
pyautogui.PAUSE = 2.5
pyautogui.dragRel(0,-70)


## read below from excel sheet
monthName = 'May-2019'
pathName = 'C:\Profiles\\awang\My Documents\Report Monthly KPI\Affinity'

saveAsExcel(analysisWindow, pathName, monthName , 'site totals')

#analysisWindow.click_input()
## drag "Site" Label down
pyautogui.PAUSE = 2.5
siteDragArea.click_input()
pyautogui.dragRel(0,60)

## drag "QA Item" Label up
qaItemHeader.click_input(button='left', double='true')
pyautogui.PAUSE = 2.5
pyautogui.moveRel(0, -20)
pyautogui.PAUSE = 2.5
pyautogui.dragRel(0,-70)

saveAsExcel(analysisWindow, pathName, monthName , 'all items')

# print(reportConfigWindow.exists())


#analysisWindow.print_control_identifiers()

