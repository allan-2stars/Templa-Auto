import pyautogui
import pandas as pd
import time
import pywinauto
from datetime import datetime

from functions.functions_utils import tm_init

## get the appliation handler from the init function
templa = tm_init()[0]
app = tm_init()[1]

print("Starting...")
# start 
mainContractsTab = templa.child_window(title='QA Items', control_type='TabItem')
mainContractsTab.click_input()
mainContractsWindow = templa.child_window(title='QA Items', control_type='Window')

templa.child_window(title="New", control_type="Button").click_input()

QAItemDetailWindow = app.window(title_re='QA Item *')
QAItemDetailWindow.wait('exists', timeout=15)
print("QA Item Window opened...")

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Add QA Items' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
# print("Reading Excel...")
for i in df.index:
    details = df['DETAILS']
    itemGroup = df['ITEM GROUP']
    #scoreCard = df['SCORE CARD']
    status = df['STATUS']

    if status[i] == "Stop":
        print("Stop here")
        break

    if status[i] == "Done":
        print(details[i]+ " is Done")
        continue

    if status[i] == "Skip":
        print(details[i]+ " is Skipped")
        continue


    # #########################
    # # add new QA Item
    # #########################

    pyautogui.PAUSE = 2.5
    pyautogui.press('tab')
    # print('click on details')
    #QAItemDetailWindow.child_window(title="Details", control_type="Text").click_input()
    pyautogui.typewrite(details[i])
    pyautogui.press('tab')
    #QAItemDetailWindow.child_window(title="Item group", control_type="Text")
    pyautogui.typewrite(itemGroup[i])
    pyautogui.press('tab')
    # QAItemDetailWindow.Save.click_input()
    pyautogui.PAUSE = 2.5
    print(details[i] +" Done.")
    QAItemDetailWindow.child_window(title="Save and new", control_type="Button").click()

QAItemDetailWindow.Close.click_input()
print("All QA Item Created now")
print("##################")
    
    
