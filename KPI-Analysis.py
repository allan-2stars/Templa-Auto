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

## get Templa ready
if (os.path.exists(r'E:\TCMS_LIVE\Client Suite')):
    templa_file = r'E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe'
    app = Application(backend='uia').connect(path=templa_file)
else:
    print('Cannot find Templa on your computer')

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')


##### defined a function for save report into specific forlder repeatively ######

def saveAsExcel(flag, window, pathName, folderName, fileName):
    ## export to excel and save
    window.child_window(title='Excel', auto_id='[Group : report Tools] Tool : Report_Excel - Index : 2 ', control_type='Button').click_input()
    saveAsWindow = window.child_window(title='Save As')
    saveAsWindow.wait('exists', timeout=15)
    print('save as window open')

    ## check if this 'save' is the same as the previouse 'save' - save in the same folder?
    if flag == 'first save round':
        addressBar = saveAsWindow.child_window(title_re='Address: *', control_type='ToolBar')
        addressBar.click_input()
        pyautogui.typewrite(pathName)
        time.sleep(1)
        pyautogui.press('enter')

        ## add a new folder if not exists
        folderNameNeeded = saveAsWindow.child_window(title=folderName, control_type='ListItem')
        if not folderNameNeeded.exists():
            print('folder NOT exists yet.')
            saveAsWindow.child_window(title='New folder', auto_id='{E44616AD-6DF1-4B94-85A4-E465AE8A19DB}', control_type='Button').click_input()
            time.sleep(2)
            pyautogui.typewrite(folderName)
            time.sleep(2)
            pyautogui.press('enter')
        ## get into the newly created folder
        folderNameNeeded.click_input(button='left', double=True)
    
    ## type the file name for saving
    saveAsWindow.child_window(title='File name:', auto_id='FileNameControlHost', control_type='ComboBox').click_input()
    pyautogui.typewrite(fileName)
    ## Save button click
    saveAsWindow.child_window(title='Save', auto_id='1', control_type='Button').click_input()
    time.sleep(2)

############### function end ########################

########################
#
# Setup Excel Sheet
#
########################
if datetime.now().month - 1 == 0:
    analysis_month = 12
    analysis_year = datetime.now().year - 1
else:
    analysis_month = datetime.now().month - 1
    analysis_year = datetime.now().year

if analysis_month == 1:
    analysis_month_text = 'Jan'
if analysis_month == 2:
    analysis_month_text = 'Feb'
if analysis_month == 3:
    analysis_month_text = 'Mar'
if analysis_month == 4:
    analysis_month_text = 'Apr'
if analysis_month == 5:
    analysis_month_text = 'May'
if analysis_month == 6:
    analysis_month_text = 'Jun'
if analysis_month == 7:
    analysis_month_text = 'Jul'
if analysis_month == 8:
    analysis_month_text = 'Aug'
if analysis_month == 9:
    analysis_month_text = 'Sep'
if analysis_month == 10:
    analysis_month_text = 'Oct'
if analysis_month == 11:
    analysis_month_text = 'Nov'
if analysis_month == 12:
    analysis_month_text = 'Dec'

monthName = analysis_month_text
yearName = str(analysis_year)


site_reallocate_sheet = 'KPI Analysis' 
df = pd.read_excel('test.xlsx', sheet_name=site_reallocate_sheet)
print('starting...')
print('analysis month: ' + monthName + ' analysis year: ' + yearName)
print('analysis month text: ' + analysis_month_text)

########################################################################
####                                                                ####
############           ANALYSIS & GENERATE REPORT          #############
## recursively generate analysis report and export to local drive ######
##
########################################################################
for i in df.index:
    reportTitle = df['TITLE']
    #monthName = df['MONTH']
    #yearName = df['YEAR']
    fileNameSiteTotals = df['FILE_NAME_SITE_TOTALS']
    fileNameAllItems = df['FILE_NAME_ALL_ITEMS']
    filePath = df['PATH']
    status = df['STATUS']

    if status[i] == 'Done' or status[i] == 'Skip':
        print(str(reportTitle[i]) + ' is Done')
        continue

    if status[i] == 'Stop':
        print('Stop here')
        break


    ## start a report with title, need open one of the report analyser first
    if i < 1:
        previouseAnalysisWindow = app.window(title=str(reportTitle[i]))
    else: 
        previouseAnalysisWindow = app.window(title=str(reportTitle[i-1]))
        print('last report is',str(reportTitle[i-1]))
    print('report now is,', str(reportTitle[i]))
    analysisWindow = app.window(title=str(reportTitle[i]))
    ## open the report selection window
    ## previouseAnalysisWindow['Select live report'].click_input()  ## too slow
    liveReportButton = previouseAnalysisWindow.child_window(title='Select live report', auto_id='[Group : report Tools] Tool : Select - Index : 5 ', control_type='Button')
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
    analysisWindow['Run report'].click_input()
    print('Data loading ...')
    
    ## Header defined
    siteHeader = analysisWindow.child_window(title='Site', control_type='ComboBox')
    siteHeader.wait('exists', timeout=280)
    ## once the report loaded, start generating...
    dragArea = analysisWindow.child_window(auto_id='GroupByBox', control_type='Group')
    qaItemHeader = analysisWindow.child_window(title='QA Item',  control_type='ComboBox')
    print('Data loaded, report generating ...')
    ## Drag area title defined
    siteDragArea = analysisWindow.child_window(title='Site', control_type='Button')
    qaItemDragArea = analysisWindow.child_window(title='QA Item', control_type='Button')

    ## drag 'Site' Label up
    siteHeader.click_input(button='left', double='true')
    time.sleep(1)
    pyautogui.moveRel(0, -20)
    time.sleep(1)
    pyautogui.dragRel(0,-70)


    ## read below from excel sheet
    folderName = monthName + '-' + yearName
    print('folder name: ' + folderName)
    
    ## save to the folder
    ## setup a flag, if the same site, no need to change the saving path
    saveAsExcel('first save round', analysisWindow, filePath[i], folderName , fileNameSiteTotals[i])

    #analysisWindow.click_input()
    ## drag 'Site' Label down
    time.sleep(1)
    siteDragArea.click_input()
    pyautogui.dragRel(0,60)

    ## drag 'QA Item' Label up
    qaItemHeader.click_input(button='left', double='true')
    time.sleep(1)
    pyautogui.moveRel(0, -20)
    time.sleep(1)
    pyautogui.dragRel(0,-70)

    ## save to the folder
    saveAsExcel('same save round', analysisWindow, filePath[i], folderName , fileNameAllItems[i])

    print(str(reportTitle[i]) + ': is Done now')
    print('###############################')
    print(' ')

