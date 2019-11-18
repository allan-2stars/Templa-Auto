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

## get Templa ready
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')


##### defined a function for save report into specific forlder repeatively ######

def saveAsExcel(window, pathName, folderName, fileName):
    ## export to excel and save
    window.child_window(title="Excel", auto_id="[Group : report Tools] Tool : Report_Excel - Index : 2 ", control_type="Button").click_input()
    saveAsWindow = window.child_window(title='Save As')
    saveAsWindow.wait('exists', timeout=15)
    print('save as window open')
    addressBar = saveAsWindow.child_window(title_re="Address: *", control_type="ToolBar")
    addressBar.click_input()
    pyautogui.typewrite(pathName)
    time.sleep(1)
    pyautogui.press('enter')
    ## add a new folder if not exists
    folderNameNeeded = saveAsWindow.child_window(title=folderName, control_type="ListItem")
    if not folderNameNeeded.exists():
        print('folder NOT exists yet.')
        saveAsWindow.child_window(title="New folder", auto_id="{E44616AD-6DF1-4B94-85A4-E465AE8A19DB}", control_type="Button").click_input()
        time.sleep(2)
        pyautogui.typewrite(folderName)
        time.sleep(2)
        pyautogui.press('enter')
    ## get into the newly created folder
    folderNameNeeded.click_input(button='left', double=True)
    ## File name type
    saveAsWindow.child_window(title="File name:", auto_id="FileNameControlHost", control_type="ComboBox").click_input()
    pyautogui.typewrite(fileName)
    ## Save button click
    saveAsWindow.child_window(title="Save", auto_id="1", control_type="Button").click_input()
    time.sleep(2)

############### function end ########################

########################
#
# Setup Excel Sheet
#
########################
site_reallocate_sheet = 'KPI Analysis' 
df = pd.read_excel('test.xlsx', sheet_name=site_reallocate_sheet)
print("starting...")


########################################################################
####                                                                ####
############           ANALYSIS & GENERATE REPORT          #############
## recursively generate analysis report and export to local drive ######
##
########################################################################
for i in df.index:
    reportTitle = df['TITLE']
    monthName = df['MONTH']
    yearName = df['YEAR']
    fileNameSiteTotals = df['FILE_NAME_SITE_TOTALS']
    fileNameAllItems = df['FILE_NAME_ALL_ITEMS']
    filePath = df['PATH']
    status = df['STATUS']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(reportTitle[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break


    ## start a report with title, need open one of the report analyser first
    if i < 1:
        previouseAnalysisWindow = app.window(title=reportTitle[i])
    else: 
        previouseAnalysisWindow = app.window(title=reportTitle[i-1])
    print('report now is,', str(reportTitle[i]))
    print('last report is',str(reportTitle[i-1]))
    analysisWindow = app.window(title=str(reportTitle[i]))
    ## open the report selection window
    ## previouseAnalysisWindow['Select live report'].click_input()  ## too slow
    liveReportButton = previouseAnalysisWindow.child_window(title="Select live report", auto_id="[Group : report Tools] Tool : Select - Index : 5 ", control_type="Button")
    liveReportButton.wait('exists',10)
    liveReportButton.click_input()
    reportConfigWindow = previouseAnalysisWindow.child_window(title='QA Analysis Report Configurations')
    reportConfigWindow.wait('exists', timeout=15)

    ## type report title 
    reportConfigWindow.window(title='Description', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(reportTitle[i]))
    pyautogui.moveRel(0, 25) 
    pyautogui.click() # open the site by double click
    reportConfigWindow.Select.click_input()

    analysisWindow.print_control_identifiers()
