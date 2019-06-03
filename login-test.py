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
##pyautogui.moveRel(-50,0)
#pyautogui.dragRel(50,0)
loginPage['PasswordEdit'].click_input()
pyautogui.typewrite('wlnde')
#pyautogui.moveRel(-50,0)
pyautogui.dragRel(-150,0)
#pyautogui.moveRel(0,100)
#pyautogui.typewrite(' ')
#loginPage['LoginButton'].click_input()


