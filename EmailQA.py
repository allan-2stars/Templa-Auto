import pyautogui
import pandas as pd
import time
import pywinauto
from datetime import datetime
from functions.functions_utils import tm_init

## get the appliation handler from the init function
templa = tm_init()[0]
app = tm_init()[1]

## start 
templa.child_window(title='QA Form List', control_type='TabItem').click_input()


mainQAListWindow = templa.child_window(title='QA Form List', control_type='Window')



########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'Email QA' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("starting...")

for i in df.index:
    siteCode = df['CODE']
    site = df['SITE']
    csm = df['CSM']
    email = df['EMAIL']
    status = df['STATUS']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(siteCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    mainQAListWindow.window(title='Site', control_type='ComboBox').click_input()
    pyautogui.typewrite(siteCode[i])
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click


    # # open qa list details dialouge window
    qaDetailWindow = app.window(title_re='QA Form Detail - *')
    qaDetailWindow.wait('exists', timeout=15)

    qaDetailWindow.child_window(title="Print/email QA", control_type="Button").click_input()


    # open qa email print dialouge window
    qaDistributeWindow = qaDetailWindow.window(title='QA form distribution')
    qaDistributeWindow.wait('exists', timeout=15)
    qaDistributeWindow.child_window(title="Email", auto_id="btnEmail", control_type="Button").click_input()

    # open qa email dialouge window
    qaEmailWindow = qaDistributeWindow.window(title='Email To')
    qaEmailWindow.wait('exists', timeout=15)
    qaEmailWindow.child_window(title="Other", control_type="RadioButton").click_input()
    pyautogui.press('tab')
    pyautogui.typewrite(email[i])
    print("send email to: " + csm[i])
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')

    # Click ok on Confirm action
    qaConfirmPopup = qaDistributeWindow.window(title='Confirm action')
    qaConfirmPopup.wait('exists', timeout=15)
    qaConfirmPopup.OK.click_input()

    # Close QA Form Detail Window
    qaDetailWindow.Close.click_input(double=True)
    print("site done: " + site[i])
    print("##########################")
    print(" ")


