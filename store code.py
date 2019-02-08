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


def clearTextBoxByDeleteKey():
   for i in range(0,18):
    pyautogui.press('del')

def clearTextBySelectAll():
    pyautogui.moveRel(30, 0)
    pyautogui.dragRel(-500, 0, 1, button='left')
    pyautogui.press('del')


#print(templa)
#def generate_data_file(t_interval, interface_name, file_name):
# start Wireshark
if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")

    

templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')
mainSitesTab = templa.child_window(title='Sites', control_type='TabItem')
#mainSitesTab.click_input()
mainSitesWindow = templa.child_window(title='Sites', control_type='Window')



site_reallocate_sheet = 'Sites Re-Allocate' 
df = pd.read_excel('test.xlsx', sheetname=site_reallocate_sheet)


for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    csm = df['CSM']
    ipad = df['IPAD']

    # click on the Code Edit Box
    mainSitesWindow.window(title='Code', control_type='ComboBox').click_input()
    pyautogui.typewrite(siteCode[i])
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    pyautogui.PAUSE = 10.5
    siteDetailWindow = app.findwindows.find_windows(title_re = 'Site Detail - *')
    #siteDetailWindow = app.window(title_re='Site Detail - *')
    analysisVersionTab = siteDetailWindow.child_window(title="Analysis versions", control_type="TabItem")
    analysisVersionTab.click_input()
    #siteDetailWindow.click_input()
    siteDetailWindow['AddButton'].click_input()
    pyautogui.PAUSE = 2.5 

    # click on Business analysis tab
    siteDetailWindow["Business analysis"].click_input() 
	# Get 2x Edit Text Box: CSM and Tablet
    csmItem = siteDetailWindow.child_window(title="CSM", control_type="DataItem")
    csmEdit = csmItem.child_window(title="Code", control_type="Edit")
    csmEdit.click_input()
    ## clear the original text
    clearTextBySelectAll()
    # Type in the new CSM Name
    pyautogui.typewrite(csm)
    pyautogui.press('tab')

    tabletItem = siteDetailWindow.child_window(title="Tablet", control_type="DataItem")
    tabletEdit =tabletItem.child_window(title="Code", control_type="Edit")
    tabletEdit.click_input()
    ## clear the original text
    clearTextBySelectAll()

    # Type in the new iPad number
    pyautogui.typewrite(ipad)
    pyautogui.press('tab')
    pyautogui.moveRel(500, 30) # move the mouse back
    pyautogui.doubleClick()

    acceptButton = siteDetailWindow.child_window(title='Accept', control_type='Button')
    acceptButton.click_input()

    siteDetailWindow["Save"].click_input()



