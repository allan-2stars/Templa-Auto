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

        ## skipped_count counts the lines of spreadsheet,
        ## see how many lines marked as "Done" or "Skip",
        ## if the total number of skipped_count equals i, which means,
        ## all line are skpped to run. Then finish the function directly.
        skipped_count = 0
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
                skipped_count = skipped_count + 1
                continue

            if status[i] == 'Skip':
                print(str(reportTitle[i]) + ' is Skipped')
                skipped_count = skipped_count + 1
                continue

            if status[i] == 'Stop':
                print('Stop here')
                ## if stopped the last Analysis window will not close
                ## due to the counter will stop counting and not reach the bottom code.
                break

            analysisWindow = app.window(title=str(reportTitle[i]))
            ## start a report with title, need open one of the report analyser first
            ## if i == 0, which mean the first line runs, and just click Run report button
            ## streight away

            if i != 0: 
                previouseAnalysisWindow = app.window(title=str(reportTitle[i-1]))
                print('last report is',str(reportTitle[i-1]))
                print('report now is,', str(reportTitle[i]))
                analysisWindow = app.window(title=str(reportTitle[i]))
                ## open the report selection window
                ## previouseAnalysisWindow['Select live report'].click_input()  ## too slow
                liveReportButton = previouseAnalysisWindow.child_window(title='Select live report', auto_id='[Group : report Tools] Tool : Select - Index : 5 ', control_type='Button')
                liveReportButton.wait('exists',20)
                liveReportButton.click_input()
                reportConfigWindow = previouseAnalysisWindow.child_window(title='QA Analysis Report Configurations')
                reportConfigWindow.wait('exists', timeout=55)

                ## type report title 
                reportConfigWindow.window(title='Description', control_type='ComboBox').click_input()
                pyautogui.typewrite(str(reportTitle[i]))
                pyautogui.moveRel(0, 25) 
                pyautogui.click() # open the site by double click
                reportConfigWindow.Select.click_input()
            
            ## Press Run report button
            analysisWindow['Run report'].click_input()

            ## calculate the loading time
            start = time.time()
            print('Data loading ...')
            
            ## Header defined
            siteHeader = analysisWindow.child_window(title='Site', control_type='ComboBox')
            siteHeader.wait('exists', timeout=280)  ## wait for the analysis result
            end = time.time()
            print('Time taken for loading the result: ', str(int(end - start)) + ' sec')
            ## once the report loaded, start generating...
            dragArea = analysisWindow.child_window(auto_id='GroupByBox', control_type='Group')
            qaItemHeader = analysisWindow.child_window(title='QA Item',  control_type='ComboBox')
            qaSiteAreaHeader = analysisWindow.child_window(title='QA Site Area',  control_type='ComboBox')
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
            save_as_Excel_analysis(window=analysisWindow, pathName=filePath[i], folderName=folderName, \
                                fileName=fileNameSiteTotals[i], flag='first time save to this folder')

            #analysisWindow.click_input()
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
                ## drag 'Site' Label to the right
                siteHeader.click_input(button='left', double='true')
                time.sleep(1)
                pyautogui.moveRel(0, -20)
                time.sleep(1)
                pyautogui.dragRel(100, 0)
            else:
                qaItemHeader.click_input(button='left', double='true')
                time.sleep(1)
                pyautogui.moveRel(0, -20)
                time.sleep(1)
                pyautogui.dragRel(0,-70)

            ## save to the folder
            save_as_Excel_analysis(window=analysisWindow, pathName=filePath[i], folderName=folderName , \
                                    fileName=fileNameAllItems[i])

            print(str(reportTitle[i]) + ': is Done now')
            print('###############################')
            print(' ')

        ## if total lines are skipped, then do nothing and finish
        ## otherwise, close analysis window
        number_of_rows = len(df.index)
        if skipped_count != number_of_rows:
            ## if this is the last item need to be analysis in Excel sheet,
            ## close the active window by press Alt+F4.
            print('closing analysis window now ...')
            keyboard.send_keys('%{F4}')
        # print('total number of rows: ', number_of_rows)
        # print('total skipped count: ', skipped_count)
        print('total skipped')
        
        ## analysisWindow.Close.click_input()
        ## analysis_window = app.window(title_re='*Monthly'), control_type='Button').Close.click_input()
        print('finishing analysis...')