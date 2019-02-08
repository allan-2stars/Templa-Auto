from subprocess import Popen
from pywinauto import Desktop
import pyautogui
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from pywinauto.application import Application
import time
import csv
import os
import sys

#print(templa)
#def generate_data_file(t_interval, interface_name, file_name):
# start Wireshark
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').start(templa_file)
    #app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")


loginPage = app['TemplaCMS  -  Login']

loginPage['Edit'].click_input()
pyautogui.typewrite('awa')
loginPage['PasswordEdit'].click_input()
pyautogui.typewrite('wlnce')
loginPage['LoginButton'].click_input()


# Error Active User Exist
errorWindow = loginPage.window(title_re="Existing*")
time.sleep(5)
#errorWindow.wait("exists",timeout=15)
print("Starting...")
if errorWindow.exists():
    print("Active user exists...")
    redCross = errorWindow.child_window(auto_id="43", control_type="Edit")
    redCross.click_input()
    pyautogui.press('y')
    errorWindow.Continue.click_input()

print("Get in Main Window...")
templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')
templa.wait("exists", timeout=15)

templa.child_window(title="Contract Admin", control_type="Group").click_input()

## Open Contract
contractsSubMenu = templa.child_window(title="Contracts", control_type="DataItem")
contractsSubMenu.click_input()
##################
## Defautl Filter
##################
contractsFilterWindow = templa.window(title_re='Contract Filter Detail -*')
# Wait filter comes out
contractsFilterWindow.wait('exists', timeout=15)
contractsFilterWindow.child_window(title="Default criteria").click_input()
contractsFilterWindow.Save.click_input()
pyautogui.PAUSE = 6.5

###########
#
# Open Sites
#
############
sitesSubMenu = templa.child_window(title="Sites",control_type="DataItem")
sitesSubMenu.click_input()
##################
## Defautl Filter
##################
sitesFilterWindow = templa.window(title_re='Site Filter Detail - *')
# Wait filter comes out
sitesFilterWindow.wait('exists', timeout=15)
sitesFilterWindow.child_window(title="Default criteria").click_input()
sitesFilterWindow.Save.click_input()
pyautogui.PAUSE = 6.5

###########
#
# Open Contacts
#
############
contactsSubMenu = templa.child_window(title="Contacts",control_type="DataItem")
contactsSubMenu.click_input()
##################
## Defautl Filter
##################
contactsFilterWindow = templa.window(title_re='Contact Filter Detail - *')
# Wait filter comes out
contactsFilterWindow.wait('exists', timeout=15)
contactsFilterWindow.child_window(title="Default criteria").click_input()
contactsFilterWindow.Save.click_input()
pyautogui.PAUSE = 6.5
