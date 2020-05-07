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



    # # If the product exists
    # # click on the "Code" Edit Box
    # mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
    # pyautogui.typewrite(str(productCode[i]))

    # click on the "Product group" Edit Box
    # print("select product group...")
    # mainProductsWindow.window(title='Product group', control_type='ComboBox').click_input()
    # pyautogui.typewrite(category[i])
    # pyautogui.moveRel(20, 25) 
    # pyautogui.click()
    # templa.child_window(title="Copy", control_type="Button").click_input()
    # print("copied one product..")





    # # click on the Code Edit Box
    # mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
    # pyautogui.typewrite(str(productCode[i]))

    # # if no product showing, then click on clear fileter button
    # # and try next one
    # productItem = mainProductsWindow.child_window(title=str(productCode[i]), control_type="DataItem")
    # if productItem.exists():
    #     # click clear button
    #     print (str(productCode[i]) + " Exist")
    #     mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
        
    # else:  
    #     print (str(productCode[i]) + " Product not exist, continue...")
    #     #

    #     #######################################
    #     #
    #     # Copy old Products
    #     #
    #     #######################################
    #     # click on other textbox first to deselect text
    #     mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()



    # indent once below

    # click on the Code Edit Box
    mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
 
    #######################################
    #
    # Setup Copied Product Code
    #
    #######################################

    # Setup Copied Product Code.
    # if not all the same setup in Excel Sheet
    # for example copy product CODE URB120
    existingCode = str(productCode[i])
    pyautogui.typewrite(existingCode)
    mainProductsWindow.child_window(title=existingCode, control_type="DataItem").click_input()





    # Click COPY to copy the product
    templa.child_window(title="Copy", control_type="Button").click_input()

    productDetailWindow = app.window(title_re='Products - *')
    productDetailWindow.wait('exists', timeout=15)

    #productDetailWindow.print_control_identifiers()
    # Type code
    productDetailWindow.child_window(auto_id="txtCode", control_type="Edit").click_input()
    pyautogui.typewrite(productCode[i])

    # just tab will select all text, no need to clear manually
    pyautogui.press('tab')
    pyautogui.typewrite(itemName[i])

    print("general info filled")
    ###################################
    # 
    # Need add Product code in the first page
    #
    ###################################
    
    # Go to Price Group, change selling price
    productDetailWindow.window(title='Price groups', control_type='TabItem').click_input()
    # find the Client Name
    priceGroupTextBox = productDetailWindow.child_window(title=clientName[i], control_type="DataItem")
    FixedPriceTextBox = priceGroupTextBox.child_window(title="Fixed price", control_type="Edit")
    FixedPriceTextBox.click_input()
    pyautogui.typewrite(str(salePrice[i]))
    print ("Sale Price is: " + str(salePrice[i]))


    # then change the cost price
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
        pyautogui.typewrite(str(productCode[i]))
        pyautogui.press('tab')
        # check prefer checkbox
        # add/change price
        #productSupplierWindow.child_window(auto_id="numUnitCost", control_type="Edit").click_input()
        pyautogui.typewrite(str(cost[i]))
        print ("Buying Price entered: " + str(cost[i]))

    else:  
        # open specific supplier item
        supplierEntry.click_input(button='left', double=True)
        productSupplierWindow = productDetailWindow.child_window(title_re='Product suppliers - *')
        productSupplierWindow.wait('exists', timeout=15)
        # add/change price
        productSupplierWindow.child_window(auto_id="numUnitCost", control_type="Edit").click_input()
        pyautogui.typewrite(str(cost[i]))
        pyautogui.keyDown('shift')
        pyautogui.press('tab')
        pyautogui.keyUp('shift')
        pyautogui.typewrite(str(productCode[i]))

    pyautogui.press('tab')

    # Save
    productSupplierWindow.Accept.click_input()
    pyautogui.PAUSE = 2.5
    productDetailWindow.Save.click_input()
    pyautogui.PAUSE = 2.5
    print (str(productCode[i]) + " is Done now")

print ("###################################")
print (" ")



    


