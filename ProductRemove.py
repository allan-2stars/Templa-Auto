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
templa.child_window(title='Product List', control_type='TabItem').click_input()
mainProductsWindow = templa.child_window(title='Product List', control_type='Window')

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Remove Product'
df = pd.read_excel('test.xlsx', sheetname=sheetLoader)

for i in df.index:
    productCode = df['PRODUCT-CODE']
    supplierCode = df['SUPPLIER-CODE']
    supplierName = df['SUPPLIER-NAME']
    status = df['STATUS']

    # supplier name is code and name combined separate with ' - '
    supplier_full_name = str(supplierCode[i]) + ' - ' + str(supplierName[i])

    if status[i] == "Done" or status[i] == "Skip":
        print(str(productCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code/Items name to Edit Box      
    mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(productCode[i]))

    # # in case the code is not long enough to narrow to one product
    # # so we need item description to narrow it down further
    # mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
    # pyautogui.typewrite(itemName[i])
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click
    print('starting...')
    print('Checking Product Code# ', productCode[i])

    # # open analysis details dialouge window
    # #siteDetailWindow = app.window(title_re='Site Detail - *')
    productDetailWindow = app.window(title_re='Products - *')
    productDetailWindow.wait('exists', timeout=15)

    # Go to supplier, only change the cost price
    productDetailWindow.window(title='Suppliers', control_type='TabItem').click_input()

    supplierEntry = productDetailWindow.child_window(title_re=supplier_full_name)
    if not supplierEntry.exists():
        productDetailWindow.Add.click_input()
        print ('supplier not exist...exit')
        
    else:  
        # open specific supplier item
        supplierEntry.click_input()
        print('supplier exists, deleting...')
        productDetailWindow["Remove"].click_input()


    # Save
    pyautogui.PAUSE = 2.5
    productDetailWindow.Save.click_input()
    pyautogui.PAUSE = 2.5
    print (supplier_full_name + ' is Done now')



    


