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

from functions.functions_utils import tm_init

## get the appliation handler from the init function
templa = tm_init()[0]
app = tm_init()[1]

print("Starting...")
## start 
mainContractsTab = templa.child_window(title='Contracts', control_type='TabItem')
mainContractsTab.click_input()
mainContractsWindow = templa.child_window(title='Contracts', control_type='Window')

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'BackToLastQA' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("Reading Excel...")
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    status = df['STATUS']
    qaType = df['QAs TYPE']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(siteCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code Edit Box
    mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(siteCode[i]))
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    print("contiune...")
    print("site code is: " + str(siteCode[i]))

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=15)


    # Go to QA tab
    contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
    ################################
    ##
    ## below conditions for in case
    ## there are multiple QA Items you have, and need to change
    ##
    # titleBox = contractDetailWindow.child_window(title="Title", auto_id="16", control_type="ComboBox")
    # isMultipleItems = False
    # if title[i] != "":
    #     titleBox.click_input()
    #     pyautogui.typewrite(title[i])
    #     isMultipleItems = True

    ################################
    # see if QA item exist

    # see if exist	    # see if exist
    qaInternalItem = contractDetailWindow.window(title='1')
    qaExternalItem = contractDetailWindow.window(title='2')
    qaExternalItemOther = contractDetailWindow.window(title='4')


    # if item exist, then see if need to change freq
    if  qaInternalItem.exists() or qaExternalItem.exists() or qaExternalItemOther.exists():
        
        contractDetailWindow.window(title='New version').click_input(double=True)

        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        time.sleep(5)

        # press the tab of QA
        contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()
        qaTemplateBox = contractDetailWindow.child_window(title="QA template", control_type="ComboBox")
        qaTemplateBox.click_input()
        QA_Type = str(qaType[i])
        if QA_Type == 'Meet and Greet':
            pyautogui.typewrite(QA_Type)
            time.sleep(2)
            qaInternalItem.wait('exists', timeout=3)
            qaInternalItem.click_input(double=True)
        else:
            qaExternalItem.click_input(double=True)
        # ################################
        
        # if existsExternalItem:
        #     qaExternalItem.click_input(double=True)
        # if existsExtItem:
        #     qaExternalItemExt.click_input(double=True)
        # if existContractItem:
        #     qaContractItem.click_input(double=True)
        # else:
        #     print ("QA Item not found, exit...")
        #     break

        print ("openning the qa item...")

        contractDetailWindow.child_window(title='Edit this effective version').click_input()
        pyautogui.PAUSE = 2.5

        contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
        contractQAWindow.wait('exists', timeout=15)


        # # If needed, un-comment below function 
        # # for change QA template at the same time.
        
        # ######################
        # #
        # # Change QA Template
        # #
        # ######################

        
        contractQAWindow.child_window(auto_id="datLastQA", control_type="Edit").click_input()
        pyautogui.hotkey('ctrl','c')

        contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('tab')

        # Save
        contractQAWindow.Accept.click_input()
        contractDetailWindow.window(title='Request approval').click_input()
        pyautogui.PAUSE = 2.5
        pyautogui.typewrite('y') ## equivilent to clicking "yes"
        print(str(siteCode[i]) + " updated now")
        #time.sleep(16)
        
    # if no qa, close it
    else:
        contractDetailWindow.Close.click_input()
        print ("No QA for this site, closed directly.")
       

    print("Done:", str(siteCode[i]))
    print("##################")
    print("                  ")
    
    
