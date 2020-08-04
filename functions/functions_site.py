import pyautogui
import pywinauto
import pandas as pd
import time
import csv
from datetime import datetime
from functions.functions_utils import tm_init


## get the appliation handler from the init function


#############################
##
## Site Reassign function
## Important Note:
##   in Templa system, 
##   make sure the CSM columu the first columu
##
#############################
def Site_Reassign():
    if tm_init() is None:
        print("Can't find Templa on your computer")
    else:
        templa = tm_init()[0]
        app = tm_init()[1]
        mainSitesTab = templa.child_window(title='Sites', control_type='TabItem')
        mainSitesTab.click_input()
        mainSitesWindow = templa.child_window(title='Sites', control_type='Window')

        ########################
        #
        # Setup Excel Sheet
        #
        ########################
        site_reallocate_sheet = 'Sites Re-Allocate' 
        df = pd.read_excel('test.xlsx', sheet_name=site_reallocate_sheet)
        print("starting...")

        for i in df.index:
            siteCode = df['CODE']
            siteName = df['SITE']
            csm = df['CSM']
            tablet = df['TABLET']
            status = df['STATUS']

            site_code_string = str(siteCode[i])
            site_name_string= str(siteName[i])
            #print("Site Name:" + siteName[i])
            #print("CSM: " + csm[i])
            #print("iPad: " + ipad[i])
            if status[i] == "Done":
                print(site_code_string + " is Done")
                continue
            if status[i] == "Skip":
                print(site_code_string + " is Skipped")
                continue

            if status[i] == "Stop":
                print("Stop here")
                break

            # click on the Code Edit Box
            mainSitesWindow.window(title='Code', control_type='ComboBox').click_input()
            pyautogui.typewrite(site_code_string)

            #####################################
            # before open the site, 
            # check if the site already set up correctly
            # check the CSM Name
            #####################################

            # mainSitesWindow.child_window(title="CSM", control_type="ComboBox").click_input()
            # pyautogui.typewrite(csm[i])

            # check if the CSM already assigned to this site
            #
            #   MUST make the CSM on the first Column
            #
            ## below function does not work
            #  regex_string_for_csm = '/' + csm[i] + '/i'

            csmExists = mainSitesWindow.child_window(title=csm[i], control_type="DataItem")
            
            
            if csmExists.exists():  
                print("site Code: " + site_code_string)
                print("site Name: " + site_name_string)
                print("Already assigned to " + csm[i])
                print("#################################")
                print(" ")
                pyautogui.moveRel(-25, 25) 
                pyautogui.click() # reset the select status

            else:
                print("CSM Different, need to change")
                print("New CSM: " + csm[i])
                pyautogui.moveRel(-25, 25) 
                pyautogui.doubleClick() # open the site by double click


                # # open analysis details dialouge window
                # #siteDetailWindow = app.window(title_re='Site Detail - *')
                siteDetailWindow = app.window(title_re='Site Detail - *')
                siteDetailWindow.wait('exists', timeout=15)
                siteDetailWindow.window(title='Analysis versions', control_type='TabItem').click_input()
                print("site name: " + site_name_string)


                ########################
                #
                #   Need to check if the month is current month
                #   if True: double click on itself
                #   if False: click Add button
                #
                #######################
                #siteDetailWindow.print_control_identifiers()
                currentYearFull = datetime.now().strftime('%Y')  # 2018
                currentMonth = datetime.now().strftime('%m') # month in number with 0 padding

                itemExist = False
                for j in range(1,32):  # loop from 1 to 31
                    day = str(j)
                    if j <= 9:
                        day = '0' + day 
            
                    titleDate= "%s.%s.%s" %(day,currentMonth,currentYearFull)
                    lastAnalysisItem = siteDetailWindow.window(title=str(titleDate))
                    # print('date of today: ' + titleDate)
                    if lastAnalysisItem.exists():
                        itemExist = True
                        break

                if itemExist: # if the current month entry exists
                    lastAnalysisItem.click_input(double=True)
                    print ("open the last item")
                else:
                    siteDetailWindow['Add'].click_input()
                    print ("add new entry")


                ## operate the site details analysis window
                siteAnalysisWindow = siteDetailWindow.child_window(title_re='Site Analysis Detail - *')
                siteAnalysisWindow.wait('exists', timeout=15)
                siteAnalysisWindow.window(title='Business analysis', control_type='TabItem').click_input()

                # change CSM and Tablet Number
                # Edit 39 = CSM, Edit 43 = Tablet
                siteAnalysisWindow.Edit39.click_input()
                # pyautogui.moveRel(60,0)
                time.sleep(1.5)
                pyautogui.dragRel(-500,0)
                pyautogui.typewrite(csm[i])
                pyautogui.press("tab")
                time.sleep(1.5)
                # print("Located now to: " + str(csm[i]))
                pyautogui.moveRel(500,15)
                pyautogui.click()
                # siteAnalysisWindow.Edit43.click_input()
                pyautogui.dragRel(-500,0)
                time.sleep(1.5)
                pyautogui.typewrite(tablet[i])
                pyautogui.press("tab")
                time.sleep(1.5)
                pyautogui.moveRel(300,20)
                pyautogui.click()
                pyautogui.press("tab")
                # press Accept button
                # Save
                siteAnalysisWindow.Accept.click_input()
                time.sleep(1.5)
                siteDetailWindow.Save.click_input()
                #siteDetailWindow.window(title='Save', control_type='Button')[0].click_input()
                time.sleep(1.5)
                print(site_code_string + ": is Done now")
                print("###############################")
                print(" ")
    