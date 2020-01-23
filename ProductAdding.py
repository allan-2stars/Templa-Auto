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
sheetLoader = 'Add Product' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("starting...")

for i in df.index:
    productCode = df['PRODUCT-CODE']
    category = df['CATEGORY']
    supplierCode = df['SUPPLIER-CODE']
    supplierCodeRe = df['SUPPLIER-CODE-RE']
    itemName = df['ITEMS']
    cost = df['COST']
    salePrice = df['SALE PRICE']
    clientName = df['CLIENT NAME']
    status = df['STATUS']
    prefer = df['PREFERRED']
    unit = df['UNIT']
    preferString = str(int(prefer[i]))

    if status[i] == "Done" or status[i] == "Skip":
        print(str(productCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break


    # trim the space on leading and tailing, in case excel sheet code un-stripped.
    productPartNo = str(productCode[i]).strip()

    # # If the product exists
    # # click on the "Code" Edit Box
    mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
    pyautogui.typewrite(productPartNo)


    # check if the Part Number is match the Product code, move Part no. on the first column
    productItem = mainProductsWindow.child_window(title=productPartNo, control_type="DataItem")
    # get category and ready for check exisitance

    productDetailWindow = app.window(title_re='Products - *')
    # if the product exists already, then directly open it.
    if productItem.exists():
        print(productPartNo + "Product Exists, open it directly")
        pyautogui.moveRel(-25, 25) # move the cursor and selec the product
        pyautogui.doubleClick() # open the existing product by double click
    # if not exists, create a new one.
    else:  
        print (productPartNo + " Product not exist, add as new...")
        #
        # Add as New Products when product not exists yet
        templa.child_window(title="New", control_type="Button").click_input()

        #######################################
        #
        # Update Products
        #
        #######################################        
        productDetailWindow.wait('exists', timeout=15)

        #productDetailWindow.print_control_identifiers()
        # Type code
        productDetailWindow.child_window(auto_id="txtCode", control_type="Edit").click_input()
        pyautogui.typewrite(productPartNo)
        productDetailWindow.child_window(auto_id="txtDescription", control_type="Edit").click_input()
        pyautogui.typewrite(itemName[i])

        # Tab 2x times to Product Category
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(category[i])
        # Tab 7x time to Purchased as Unit
        pyautogui.press('tab')
        pyautogui.press('tab')   
        pyautogui.press('tab')
        pyautogui.press('tab')    
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(unit[i])
        # Then keep going to all Unit
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(unit[i])
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(unit[i])
        pyautogui.press('tab')

        print("general info filled")

        ###################################
        # 
        # Need add Product code in the first page
        #
        ###################################
        
    # Go to Price Group, change selling price
    # if no need to setup sale price, then clientname will be "na"
    hasSalePrice = str(clientName[i]) != 'nan'
    print(str(clientName[i]))
    print ("has client? " + str(hasSalePrice))
    if hasSalePrice:
        productDetailWindow.window(title='Price groups', control_type='TabItem').click_input()
        # find the Client Name
        priceGroupTextBox = productDetailWindow.child_window(title=clientName[i], control_type="DataItem")
        FixedPriceTextBox = priceGroupTextBox.child_window(title="Fixed price", control_type="Edit")
        FixedPriceTextBox.click_input()
        pyautogui.typewrite(str(salePrice[i]))
        print ("Sale Price is: " + str(salePrice[i]))

    ###################################
    # 
    # Need add Product code in the first page
    #
    ###################################

    # Go to supplier, only change the cost price
    productDetailWindow.window(title='Suppliers', control_type='TabItem').click_input()
    supplierEntry = productDetailWindow.child_window(title_re=supplierCodeRe[i])
    if not supplierEntry.exists():
        productDetailWindow.Add.click_input()
        print ("supplier not exist in the list")
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
        if str(isChecked) != preferString:
            preferredCheckbox.toggle()
        # you can also use tab tab to go down
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(productPartNo)
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
    print (productPartNo + " is Done now")

    print ("###################################")
    print (" ")



    


