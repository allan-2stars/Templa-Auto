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
sheetLoader = 'Product Checking' 
df = pd.read_excel('test.xlsx', sheetname=sheetLoader)
print("starting...")

for i in df.index:
    productCode = df['PRODUCT-CODE'] 
    status = df['STATUS']


    if status[i] == "Done":
        print(str(productCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    #######################################
    #
    # input code to check
    #
    #######################################
        # click on the Code Edit Box

    productPartNo = str(productCode[i]).strip()
    mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
    pyautogui.typewrite(productPartNo)

    # if no product showing, then click on clear fileter button
    # and try next one
    productItem = mainProductsWindow.child_window(title=productPartNo, control_type="DataItem")
    if productItem.exists():
        # click clear button
        print (productPartNo + " Exist")
    else:  
        print (productPartNo + " Not Not Exist")
        #
    mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()



    


