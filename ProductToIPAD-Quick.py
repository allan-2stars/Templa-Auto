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


storeTemplateWindow = app.top_window()
storeTemplateWindow.wait('exists', timeout=10)

print("product adding...")

#############################################
#############################################
#
# ADD PRODUCT TO TEMPLATE LIST AND TO IPAD

# Manually Open the Add multiple Window on your Templa
# Then you are ready, the programme will add code directly.
#
#############################################

print(str(storeTemplateWindow.exists()))



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
    status = df['STATUS']


    # strip product code first
    productPartNo = str(productCode[i]).strip()

    if status[i] == "Done":
        print(productPartNo + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break
    print ("code is" + productPartNo)
    storeTemplateWindow.child_window(title="Code", control_type="ComboBox").click_input()

    # click on the Code Edit Box

    pyautogui.typewrite(productPartNo)
    pyautogui.moveRel(0, 25)
    pyautogui.PAUSE = 0.5
    pyautogui.doubleClick() # open the site by double click
  

    print ("current row is: " + str(i))

   
    # productsSelectWindow.Close.click_input()
    # pyautogui.PAUSE = 2.5
    # storeTemplateWindow.Save.click_input()
    # pyautogui.PAUSE = 2.5
    # print (str(templateName[i]) + " is Done now")
    # print ("#########################")
    # print (" ")


print ("----Done----")

    


