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

#print(templa)
#def generate_data_file(t_interval, interface_name, file_name):
# start Wireshark
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')

## start 
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

    #print("Site Name:" + siteName[i])
    #print("CSM: " + csm[i])
    #print("iPad: " + ipad[i])
    if status[i] == "Done" or status[i] == "Skip":
        print(str(siteCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code Edit Box
    mainSitesWindow.window(title='Code', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(siteCode[i]))

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
    csmExists = mainSitesWindow.child_window(title=str(csm[i]), control_type="DataItem")
    
    if csmExists.exists():  
        print("site Code: " + str(siteCode[i]))
        print("site Name: " + siteName[i])
        print("Already assigned to " + csm[i])
        print("#################################")
        print(" ")
        pyautogui.moveRel(-25, 25) 
        pyautogui.click() # reset the select status

    else:
        print("CSM Different, need to change")
        pyautogui.moveRel(-25, 25) 
        pyautogui.doubleClick() # open the site by double click


        # # open analysis details dialouge window
        # #siteDetailWindow = app.window(title_re='Site Detail - *')
        siteDetailWindow = app.window(title_re='Site Detail - *')
        siteDetailWindow.wait('exists', timeout=15)
        siteDetailWindow.window(title='Analysis versions', control_type='TabItem').click_input()
        print("site name: " + str(siteName[i]))


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
            titleDate= "%s/%s/%s" %(j,currentMonth,currentYearFull)
            lastAnalysisItem = siteDetailWindow.window(title=str(titleDate))
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
        pyautogui.PAUSE = 1.5
        pyautogui.dragRel(-500,0)
        pyautogui.typewrite(str(csm[i]))
        pyautogui.press("tab")
        pyautogui.PAUSE = 1.5
        # print("Located now to: " + str(csm[i]))
        pyautogui.moveRel(500,15)
        pyautogui.click()
        # siteAnalysisWindow.Edit43.click_input()
        pyautogui.dragRel(-500,0)
        pyautogui.PAUSE = 1.5
        pyautogui.typewrite(tablet[i])
        pyautogui.press("tab")
        pyautogui.PAUSE = 1.5
        pyautogui.moveRel(300,20)
        pyautogui.click()
        pyautogui.press("tab")
        # press Accept button
        # Save
        siteAnalysisWindow.Accept.click_input()
        siteDetailWindow.Save.click_input()
        pyautogui.PAUSE = 1.5
        print(str(siteCode[i]) + ": is Done now")
        print("###############################")
        print(" ")

