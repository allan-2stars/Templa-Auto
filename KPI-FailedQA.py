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
    print('Can not find Templa on your computer')

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')



##### defined a function for save report into specific forlder repeatively ######

def saveAsExcel(window, pathName, folderName, fileName):
    ## export to excel and save
    window.child_window(title='Excel', control_type='Button').click_input()
    saveAsWindow = window.child_window(title='Save As')
    saveAsWindow.wait('exists', timeout=15)
    print('save as window open')
    addressBar = saveAsWindow.child_window(title_re='Address: *', control_type='ToolBar')
    addressBar.click_input()
    pyautogui.typewrite(pathName)
    time.sleep(1)
    pyautogui.press('enter')
    ## add a new folder if not exists
    ## check case
    upperCaseFolderName = folderName.upper()
    lowerCaseFolderName = folderName.lower()
    titleCaseFolderName = folderName.title()
    upperCaseFolder = saveAsWindow.child_window(title=upperCaseFolderName, control_type='ListItem')
    lowerCaseFolder = saveAsWindow.child_window(title=lowerCaseFolderName, control_type='ListItem')
    titleCaseFolder = saveAsWindow.child_window(title=titleCaseFolderName, control_type='ListItem')

    ## default folder name is Title Case Folder
    folderNameNeeded = titleCaseFolder

    if titleCaseFolder.exists():
        print('title cased folder Exist Already')
        folderNameNeeded = titleCaseFolder
    elif upperCaseFolder.exists():
        print('upper cased folder Exist Already')
        folderNameNeeded = upperCaseFolder
    elif lowerCaseFolder.exists():
        print('lower cased folder Exist Already')
        folderNameNeeded = lowerCaseFolder
    else:
        print('folder NOT exists yet.')
        saveAsWindow.child_window(title='New folder', control_type='Button').click_input()
        time.sleep(2)
        pyautogui.typewrite(titleCaseFolderName)
        time.sleep(2)
        pyautogui.press('enter')

        
    ## get into the newly created folder
    folderNameNeeded.click_input(button='left', double=True)
    ## File name type
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
Work_Sheet = 'KPI QA Completed Items' 
df = pd.read_excel('test.xlsx', sheet_name=Work_Sheet)

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

lastday_analysis_month = str(calendar.monthrange(analysis_year, analysis_month)[1])

dateStartString = '01' + str(analysis_month) + yearName
dateEndString = lastday_analysis_month + str(analysis_month) + yearName

print('starting...')
print('analysis month:' + monthName + 'analysis year: ' + yearName)
print('analysis month text:' + analysis_month_text)

for i in df.index:
    constracts = df['CONTRACTS']
    siteName = df['SITE NAME']
    site = df['SITE']
    client = df['CLIENT']
    template = df['TEMPLATE']
    filePath = df['PATH']
    fileName = df['FILE_NAME_FAILED_QA_ITEMS']
    status = df['STATUS']

    useContracts = df['USE CONTRACTS']
    useSite = df['USE SITE']
    useClient = df['USE CLIENT']
    useTemplate = df['USE TEMPLATE']

    if status[i] == 'Done':
        print(str(siteName[i]) + ' is Done')
        continue

    if status[i] == 'Skip':
        print(str(siteName[i]) + ' is Skipped')
        continue

    if status[i] == 'Stop':
        print('Stop here')
        break
        
    print(' ')

    ## start 
    completedQAWindow = templa.child_window(title='Completed QA Items', control_type='TabItem')
    completedQAWindow.click_input()

    templa.child_window(title='Change filter', control_type='Button').click_input()
    filterWindow = templa.window(title_re='QA Completed Item Filter Detail - *')
    filterWindow.wait('exists', timeout=15)

    ## default filter
    print('Default the Filter.')
    filterWindow.child_window(title='Default criteria').click_input()

    if useTemplate[i] == 'Yes':
        print('Use Template')
        ## Use Template Filter
        filterWindow.child_window(auto_id='cslQATemplate', control_type='Pane').click_input()
        pyautogui.typewrite(str(int(template[i])))

        pyautogui.press('tab')

    # ###############################################
    # ############                    ###############
    # ############  Basic Filtering   ###############


    ## filter on date range of audited date
    filterWindow.child_window(auto_id='datAuditDateFrom', control_type='Edit').click_input()
    pyautogui.typewrite(dateStartString)
    ##filterWindow.child_window(auto_id='datAuditDateTo', control_type='Edit').click_input()
    pyautogui.press('tab')
    pyautogui.typewrite(dateEndString)


    ## if the site is Special case, use below
    if str(siteName[i]) == 'DAWR' or str(siteName[i]) == 'PMC':
        print('Ignore the failed Items')
        # pyautogui.press('right')
        # pyautogui.press('right')
        # pyautogui.press('space')
    # ## ## click on Failed Items button to YES
    else:
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('right')
        pyautogui.press('space')


    # ###########   End of Basic Filtering    #########
    # #################################################


    ## change the site filters criteria
    siteFilterCriteria = filterWindow.child_window(title='Site filtering criteria', control_type='TabItem')
    siteFilterCriteria.click_input()


    if useContracts[i] == 'Yes':
        ## Use Contracts filter
        print('Use Contracts')
        filterWindow.child_window(title='Contracts', auto_id='5', control_type='DataItem').click_input()
        pyautogui.typewrite(str(constracts[i]))
        pyautogui.press('tab')

    if useSite[i] == 'Yes':
        ## Use Site Filter
        print('Use Site')
        filterWindow.child_window(auto_id='cslSite', control_type='Pane').click_input()
        pyautogui.typewrite(str(site[i]))
        pyautogui.press('tab')

    if useClient[i] == 'Yes':
        ## Use Client Filter
        print('Use Client')
        filterWindow.child_window(auto_id='cslClient', control_type='Pane').click_input()
        pyautogui.typewrite(str(client[i]))
        pyautogui.press('tab')

    ## check the other tab filtering
    if siteName[i] == 'Redcape':    
        protertyFilterCriteria = filterWindow.child_window(title='Property filtering criteria', control_type='TabItem')
        protertyFilterCriteria.click_input()
        ## get the handle of Group filter
        groupItem = filterWindow.child_window(title='Group', auto_id='2', control_type='DataItem')
        matchTypeSection = groupItem.child_window(title='Match type', auto_id='3', control_type='ComboBox')
        valueSection = groupItem.child_window(title='Value', auto_id='1', control_type='ComboBox')

        ## click and change the filter Equal to ...
        matchTypeSection.click_input()
        pyautogui.typewrite('e')  # e, for Equal to
        pyautogui.press('tab')
        time.sleep(2)
        valueSection.click_input()
        pyautogui.typewrite('r') # filter the site name
        pyautogui.press('tab')


    # ## Save the filter
    print('Saving the filter ...')
    filterWindow.Save.click_input()
    
    #siteDescriptionTab = completedQAWindow.child_window(title='Site description', control_type='DataItem')
    mainCompletedWindow = templa.child_window(title='Completed QA Items', control_type='Window')
    csmWindow = mainCompletedWindow.child_window(title='CSM', auto_id='56', control_type='ComboBox')
    csmWindow.wait('exists', 180)

    templa.child_window(title='Select format', control_type='Button').click_input()
    filterFormatsWindow = templa.window(title='Filtered List Formats')
    filterFormatsWindow.wait('exists', timeout=15)
    ## type the format name
    filterFormatsWindow.window(title='Description', control_type='ComboBox').click_input()

    ## if the site is Special case, use below
    if str(siteName[i]) == 'DAWR' or str(siteName[i]) == 'PMC':
        pyautogui.typewrite(str(siteName[i]))
        pyautogui.moveRel(-25, 25) 
        pyautogui.doubleClick() # apply the format
    else:
        pyautogui.typewrite('Standard Format')
        pyautogui.moveRel(-25, 25) 
        pyautogui.doubleClick() # apply the format
        


    ## read below from excel sheet
    folderName = monthName + '-' + yearName
    ##
    print('Ready to Export to Excel File ...')
    saveAsExcel(templa, filePath[i], folderName , fileName[i])
    print(str(siteName[i]) + ' is Done.')
    print('#######################')
    print(' ')




