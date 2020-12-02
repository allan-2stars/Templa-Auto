import pyautogui
import pandas as pd
import time
from datetime import datetime
import calendar
import csv
import pywinauto
from pywinauto import keyboard
from datetime import datetime
from functions.functions_utils import tm_init
from functions.functions_utils import save_as_Excel_analysis


def KPI_Analysis():
    if tm_init() is None:
        print("Can't find Templa on your computer")
    else:
        templa = tm_init()[0]
        app = tm_init()[1]

        ########################
        #
        # Setup Excel Sheet
        #
        ########################
        analysis_month = datetime.now().month - 1
        if analysis_month == 0:
            analysis_month = 12
            analysis_year = datetime.now().year - 1
        else:
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

            if status[i] == 'Done':
                print(str(reportTitle[i]) + ' is Done')
                continue

            if status[i] == 'Skip':
                print(str(reportTitle[i]) + ' is Skipped')
                continue

            if status[i] == 'Stop':
                print('Stop here')
                ## if stopped the last Analysis window will not close
                ## due to the counter will stop counting and not reach the bottom code.
                break
    
            analysis_window = app.window(title_re='.*Monthly', control_type='Window')
          
            ## open the report selection window
            
            liveReportButton = analysis_window.child_window(title='Select live report', auto_id='[Group : report Tools] Tool : Select - Index : 5 ', control_type='Button').click_input()
            report_config_window = analysis_window.child_window(title='QA Analysis Report Configurations')
            report_config_window.wait('exists', timeout=55)

            ## type report title 
            report_config_window.window(title='Description', control_type='ComboBox').click_input()
            pyautogui.typewrite(str(reportTitle[i]))
            pyautogui.moveRel(0, 25) 
            pyautogui.click() # open the site by double click
            analysis_window.Select.click_input()
            
            ## Press Run report button
            analysis_window['Run report'].click_input()

            ## calculate the loading time
            start = time.time()
            print('Data loading ...')
            
            ## Header defined
            siteHeader = analysis_window.child_window(title='Site', control_type='ComboBox')
            siteHeader.wait('exists', timeout=280)  ## wait for the analysis result
            end = time.time()
            print('Time taken for loading the result: ', str(int(end - start)) + ' sec')
            ## once the report loaded, start generating...
            dragArea = analysis_window.child_window(auto_id='GroupByBox', control_type='Group')
            qaItemHeader = analysis_window.child_window(title='QA Item',  control_type='ComboBox')
            qaSiteAreaHeader = analysis_window.child_window(title='QA Site Area',  control_type='ComboBox')
            print('Data loaded, report generating ...')
            ## Drag area title defined
            siteDragArea = analysis_window.child_window(title='Site', control_type='Button')
            qaItemDragArea = analysis_window.child_window(title='QA Item', control_type='Button')

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
            save_as_Excel_analysis(window=analysis_window, pathName=filePath[i], folderName=folderName, \
                                fileName=fileNameSiteTotals[i], flag='first time save to this folder')

            #analysis_window.click_input()
            ## drag 'Site' or other Label down back
            time.sleep(1)

            siteDragArea.click_input()
            pyautogui.dragRel(50, 60)

            ## drag 'QA Item' Label up
            
            if reportTitle[i] == "DAWR Monthly" or reportTitle[i] == "PMC Monthly" or reportTitle[i] == "TK MAXX Monthly" :
                qaSiteAreaHeader.click_input()
                time.sleep(1)
                pyautogui.moveRel(0, -20)
                time.sleep(1)
                pyautogui.dragRel(0,-70)
                time.sleep(1)
                ## drag 'QA item' Label to the left
                qaItemHeader.click_input(button='left', double='true')
                pyautogui.moveRel(0, -20)
                ## click on Site label to sort it
                print('sort by Site name ...')
                pyautogui.click()
            else:
                qaItemHeader.click_input(button='left', double='true')
                time.sleep(1)
                pyautogui.moveRel(0, -20)
                time.sleep(1)
                pyautogui.dragRel(0,-70)

            ## save to the folder
            save_as_Excel_analysis(window=analysis_window, pathName=filePath[i], folderName=folderName , \
                                    fileName=fileNameAllItems[i])

            print(str(reportTitle[i]) + ': is Done now')
            print('###############################')
            print(' ')

        ## after all analysed, close the analysis window, by pressing Alt+ F4
        keyboard.send_keys('%{F4}')
        print('finishing analysis...')