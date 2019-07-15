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


if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')

## start 
templa.child_window(title='Stores Templates', control_type='TabItem').click_input()
mainStoresTemplatesWindow = templa.child_window(title='Stores Templates', control_type='Window')

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Product-IPAD' 
df = pd.read_excel('test.xlsx', sheetname=sheetLoader)
print("starting...")

for i in df.index:
    productCode = df['PRODUCT-CODE']
    templateName = df['TEMPLATE-NAME']
    status = df['STATUS']


    if status[i] == "Done" or status[i] == "Skip":
        print(str(productCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Description Edit Box
    mainStoresTemplatesWindow.child_window(title="Description", control_type="ComboBox").click_input()
    print("printing description for template...")
    pyautogui.typewrite(templateName[i])

    # # in case the code is not long enough to narrow to one product
    # # so we need item description to narrow it down further
    # mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
    # pyautogui.typewrite(itemName[i])
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click


    storeTemplateWindow = app.top_window()
    storeTemplateWindow.wait('exists', timeout=10)

    print("product adding...")

    #############################################
    #############################################
    #
    # ADD PRODUCT TO TEMPLATE LIST AND TO IPAD
    #
    #############################################

    print(str(storeTemplateWindow.exists()))
    ## click on the tab of Products 
    app.top_window().child_window(title="Products", control_type="TabItem").click_input()
    app.top_window().child_window(title='Add multiple', control_type='Button').click_input()
    
    # Products Select Window with products List 
    productsSelectWindow = app.top_window().window(title='Products',control_type="Window")
    #app.top_window().child_window(title="Code", control_type="ComboBox").click_input()

    app.top_window().print_control_identifiers()


    # while  templateName[i] == templateName[i-1] and i <= df.index:
    #     # click on the Code Edit Box
    #     productsSelectWindow.window(title='Code', control_type='ComboBox').click_input()
    #     pyautogui.typewrite(str(productCode[i]))
    #     pyautogui.moveRel(0, 25) 
    #     pyautogui.doubleClick() # open the site by double click
    #     i= i + 1
    #     print ("i now is: " + i)

   
    # productsSelectWindow.Close.click_input()
    # pyautogui.PAUSE = 2.5
    # storeTemplateWindow.Save.click_input()
    # pyautogui.PAUSE = 2.5
    # print (str(templateName[i]) + " is Done now")
    # print ("#########################")
    # print (" ")




    


