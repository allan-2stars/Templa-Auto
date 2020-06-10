import os
from pywinauto import Application
import pyautogui
import calendar
import time
from datetime import datetime, timedelta, date
import dateutil.relativedelta as relativedelta

def tm_init():
    if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
        templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
        app = Application(backend='uia').connect(path=templa_file)
        templa = app.window(title_re='TemplaCMS*')
        return [templa, app]
    else:
        return None


def date_range(relativeMonthNumber):
    month_relative_of_today = date.today() + relativedelta.relativedelta(months=relativeMonthNumber)
    month_number_relative = month_relative_of_today.strftime("%m") # type of String


    current_year = datetime.now().year
    lastday_of_relative_month = calendar.monthrange(current_year, int(month_number_relative))[1]

    date_start_string = '01' + month_number_relative + str(current_year)
    date_end_string = str(lastday_of_relative_month) + month_number_relative + str(current_year)
    print(date_start_string)
    print(date_end_string)

    # return two values in a list
    return [date_start_string, date_end_string]


##### defined a function for save report into specific forlder repeatively ######

def save_as_Excel_analysis(**kwargs):
    ## export to excel and save
    ## get kwargs into variables, put to None as default
    window = kwargs.get('window', None)
    pathName = kwargs.get('pathName', None)
    folderName = kwargs.get('folderName', None)
    fileName = kwargs.get('fileName', None)
    flag = kwargs.get('flag', None)

    window.child_window(title='Excel', control_type='Button').click_input()
    saveAsWindow = window.child_window(title='Save As')
    saveAsWindow.wait('exists', timeout=15)
    print('save as window open')
    ## flag = kwargs.get('flag', None)
    ## check if this 'save' is the same as the previouse 'save' - save in the same folder?
    if flag == 'first time save to this folder':
        print('type in Address bar to the correct directory ...')
        saveAsWindow.child_window(title_re='Address: *', control_type='ToolBar').click_input()
        pyautogui.typewrite(pathName)
        time.sleep(2)
        pyautogui.press('enter')
        ## check case
        upperCaseFolderName = folderName.upper()
        lowerCaseFolderName = folderName.lower()
        titleCaseFolderName = folderName.title()
        upperCaseFolder = saveAsWindow.child_window(title=upperCaseFolderName, control_type='ListItem')
        lowerCaseFolder = saveAsWindow.child_window(title=lowerCaseFolderName, control_type='ListItem')
        titleCaseFolder = saveAsWindow.child_window(title=titleCaseFolderName, control_type='ListItem')
        
        if  titleCaseFolder.exists():
            print('title cased folder Exist Already')
            folderNameNeeded = titleCaseFolder
        elif upperCaseFolder.exists():
            print('upper cased folder Exist Already')
            folderNameNeeded = upperCaseFolder
        elif lowerCaseFolder.exists():
            print('lower cased folder Exist Already')
            folderNameNeeded = lowerCaseFolder
        else:  ## folder not exists
            print('folder NOT exists yet. Creating ...')
            saveAsWindow.child_window(title='New folder', control_type='Button').click_input()
            time.sleep(2)
            pyautogui.typewrite(titleCaseFolderName)
            time.sleep(2)
            pyautogui.press('enter')                
        ## get into the newly created folder
        folderNameNeeded.click_input(button='left', double=True)
        saveAsWindow.child_window(title='File name:', auto_id='FileNameControlHost', control_type='ComboBox').click_input()
    
    ## directly type the file name for saving
    pyautogui.typewrite(fileName)
    ## Save button click
    saveAsWindow.child_window(title='Save', auto_id='1', control_type='Button').click_input()
    ## press 'y' for yes to overwrite the file if asked.
    ## for now there is no condition detect for this overwirte warning.
    ## just press 'y' anyway for now.
    pyautogui.press('y')
    ## wait for seconds to go next round
    time.sleep(2)

############### function end ########################
