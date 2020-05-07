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
sheetLoader = 'Change Product Name' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("starting...")

for i in df.index:
    product_code = df['PRODUCT-CODE']
    product_name = df['ITEMS']
    status = df['STATUS']


    if status[i] == "Done" or status[i] == "Skip":
        print(str(product_code[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break


    # trim the space on leading and tailing, in case excel sheet code un-stripped.
    product_code_text = str(product_code[i]).strip()
    product_name_text = str(product_name[i]).strip()

    # # If the product exists
    # # click on the "Code" Edit Box
    mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
    pyautogui.typewrite(product_code_text)

    ## in case product not found, after 'Tab' key pressed,
    ## next time when type write the code, 
    ## will replace the last enteried text
    pyautogui.press('tab') 

    # check if the Part Number is match the Product code, move Part no. on the first column
    productItem = mainProductsWindow.child_window(title=product_code_text, control_type="DataItem")
    # get category and ready for check exisitance

    productDetailWindow = app.window(title_re='Products - *')
    # if the product exists already, then directly open it.
    if productItem.exists():
        print('Product ' + product_code_text + ' Exists, open it now ...')
        pyautogui.moveRel(-25, 25) # move the cursor and selec the product
        pyautogui.doubleClick() # open the existing product by double click
    
        #######################################
        #
        # Update Products
        #
        #######################################        
        productDetailWindow.wait('exists', timeout=15)
        pyautogui.press('tab')
        ## productDetailWindow.child_window(auto_id="txtDescription", control_type="Edit").click_input()
        pyautogui.typewrite(product_name_text)

        print("Product name/description changed ...")

        productDetailWindow.Save.click_input()
        pyautogui.PAUSE = 2.5
        print (product_code_text + " product name updated.")
    else:
        print('-----------------------------------------------------------')
        print('-- Warning! Product ' + product_code_text + ' not found, please check !!! --')
        print('-----------------------------------------------------------')
        continue

    print ("###################################")
    print (" ")



    


