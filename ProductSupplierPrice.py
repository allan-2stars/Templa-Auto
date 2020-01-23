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
sheetLoader = 'Product Cost'
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)

for i in df.index:
    productCode = df['PRODUCT-CODE']
    supplierCode = df['SUPPLIER-CODE']
    supplierCodeRe = df['SUPPLIER-CODE-RE']
    itemName = df['ITEMS']
    cost = df['COST']
    status = df['STATUS']
    prefer = df['PREFERRED']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(productCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code/Items name to Edit Box
    if sheetLoader == "Urbanest NT Price"
        mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
        pyautogui.typewrite(str(itemName[i]))
    else        
        mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
        pyautogui.typewrite(str(productCode[i]))

    # # in case the code is not long enough to narrow to one product
    # # so we need item description to narrow it down further
    # mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
    # pyautogui.typewrite(itemName[i])
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click
    print("starting...")

    # # open analysis details dialouge window
    # #siteDetailWindow = app.window(title_re='Site Detail - *')
    productDetailWindow = app.window(title_re='Products - *')
    productDetailWindow.wait('exists', timeout=15)

    # Go to supplier, only change the cost price
    productDetailWindow.window(title='Suppliers', control_type='TabItem').click_input()
    supplierEntry = productDetailWindow.child_window(title_re=supplierCodeRe[i])
    if not supplierEntry.exists():
        productDetailWindow.Add.click_input()
        print ("not exist")
        # open new supplier detail window
        productSupplierWindow = productDetailWindow.child_window(title_re='Product suppliers - *')
        productSupplierWindow.wait('exists', timeout=15)
        # add supplier name by code
        # the supplier text box is focused by default
        print ("add supplier")
        pyautogui.typewrite(supplierCode[i])
        pyautogui.press('tab')
        preferredCheckbox = productSupplierWindow.child_window(auto_id="chkIsPreferredSupplier", control_type="CheckBox")
        isChecked = preferredCheckbox.get_toggle_state()

        # check if match with Excel sheet data
        if str(isChecked) != str(prefer[i]):
            preferredCheckbox.toggle()
        # you can also use tab tab to go down
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(productCode[i])
        pyautogui.press('tab')
        # check prefer checkbox
        # add/change price
        #productSupplierWindow.child_window(auto_id="numUnitCost", control_type="Edit").click_input()
        pyautogui.typewrite(str(cost[i]))
        #pyautogui.press('tab')
    else:  
        # open specific supplier item
        supplierEntry.click_input(button='left', double=True)
        productSupplierWindow = productDetailWindow.child_window(title_re='Product suppliers - *')
        productSupplierWindow.wait('exists', timeout=15)
        # add/change price
        productSupplierWindow.child_window(auto_id="numUnitCost", control_type="Edit").click_input()
        pyautogui.typewrite(str(cost[i]))
        pyautogui.press('tab')

    # Save
    productSupplierWindow.Accept.click_input()
    pyautogui.PAUSE = 2.5
    productDetailWindow.Save.click_input()
    pyautogui.PAUSE = 2.5
    print (str(itemName[i]) + " is Done now")



    


