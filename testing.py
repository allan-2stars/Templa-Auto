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

# templa.child_window(title='Product List', control_type='TabItem').click_input()
# mainProductsWindow = templa.child_window(title='Product List', control_type='Window')

print("starting...")

# mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
# print("printing description for template...")
# pyautogui.typewrite("")




# #######################################
# #
# # Copy old Products
# #
# #######################################
# # click on other textbox first to deselect text
# mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
# # click on the Code Edit Box
# mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()


# # for example copy product CODE URB120
# existingCode = "URB120"
# pyautogui.typewrite(existingCode)
# mainProductsWindow.child_window(title=existingCode, control_type="DataItem").click_input()

# templa.child_window(title="Copy", control_type="Button").click_input()

productDetailWindow = app.window(title_re='QA Item - *')
# productDetailWindow.wait('exists', timeout=15)

# #productDetailWindow.print_control_identifiers()
# # Type code
# productDetailWindow.child_window(auto_id="txtCode", control_type="Edit").click_input()
# pyautogui.typewrite("URB2000")

# # just tab will select all text, no need to clear manually
# pyautogui.press('tab')
# pyautogui.typewrite("test name")

print("general info filled")
 # Go to Price Group, change selling price
# productDetailWindow.window(title='Price groups', control_type='TabItem').click_input()

# priceGroupTextBox = productDetailWindow.child_window(title="Urbanest Room Cleaning", control_type="DataItem")
# FixedPriceTextBox = priceGroupTextBox.child_window(title="Fixed price", control_type="Edit")
# FixedPriceTextBox.click_input()
# pyautogui.moveRel(20, 0)
# pyautogui.click()



productDetailWindow.print_control_identifiers()




    


