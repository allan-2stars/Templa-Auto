import pyautogui
import pandas as pd
import time
import pywinauto
from datetime import datetime

## if you want to save output to file
## uncomment below

# import sys
# f = open('templa-output', 'w')
# sys.stdout = f

### uncomment above

from functions.functions_utils import tm_init

## get the appliation handler from the init function
templa = tm_init()[0]
app = tm_init()[1]

print("Starting...")
## start 
mainContractsTab = templa.child_window(title='Contracts', control_type='TabItem')
mainContractsTab.click_input()
mainContractsWindow = templa.child_window(title='Contracts', control_type='Window')

########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'BackToLastWorkbill' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
for i in df.index:
    siteCode = df['CODE']
    siteName = df['SITE']
    status = df['STATUS']
    workbillCost = df['COST']

    if status[i] == "Same Contract":
        print(str(siteCode[i]) + " - Same Contract")
        continue

    if status[i] == "Skip":
        print(str(siteCode[i]) + " is Skipped")
        continue

    if status[i] == "Done":
        print(str(siteCode[i]) + " is Done")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    # click on the Code Edit Box
    mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
    pyautogui.typewrite(str(siteCode[i]))
    pyautogui.moveRel(0, 25) 
    pyautogui.doubleClick() # open the site by double click

    print("contiune...")
    print("site code is: " + str(siteCode[i]))

    # # open analysis details dialouge window
    contractDetailWindow = app.window(title_re='Contract - *')
    contractDetailWindow.wait('exists', timeout=25)
    
    ## New version button click
    
    contractDetailWindow.window(title='New version').click_input()
    pyautogui.typewrite('y') ## equivilent to clicking "yes"

    # Go to QA tab
    contractDetailWindow.wait('exists', timeout=25)
    contractDetailWindow.child_window(title='Workbills', control_type='TabItem').click_input()
    
    while True:

        ## if next line is not the same section, 
        ## jump out of the while loop, need to add to a new section
        contractDetailWindow.child_window(title="Cost", control_type="ComboBox").click_input()

        pyautogui.typewrite(str(workbillCost[i]))
        pyautogui.moveRel(0, 20) 
        pyautogui.doubleClick() # open the site by double click

        contractDetailWindow.child_window(title='Edit this effective version').click_input()
        time.sleep(2.5)

        ## open contract workbill details window
        contractWorkbillWindow = contractDetailWindow.window(title_re='Contract Workbill*')
        contractWorkbillWindow.wait('exists', timeout=15)
        # contractWorkbillWindow.print_control_identifiers()

        contractWorkbillWindow.child_window(auto_id="datLastWorkbill", control_type="Edit").click_input()
        pyautogui.hotkey('ctrl','c')

        # contractWorkbillWindow.child_window(auto_id="datNextWorkbill", control_type="Edit").click_input() # next qa edit box
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('tab')

        # Save the changes
        contractWorkbillWindow.Accept.click_input()
        ## next line still int the same section, Go to next line
        print(str(siteCode[i]) + ' ' + str(workbillCost[i]) + " updated now")
        try: ## check if next line exists in spreadsheet
            next_siteCode = siteCode[i+1]
        except: ## not exist, then must be the last item already
            next_siteCode = 'NO CODE AVALIABLE!!!'

        if siteCode[i] != next_siteCode:
            contractDetailWindow.window(title='Request approval').click_input()
            time.sleep(2.5)
            pyautogui.typewrite('y') ## equivilent to clicking "yes"
            print(str(siteName[i]) + " Done.")
            print('-----------------------')
            print('')
            break  # note, once break i will not plus one
        i = i + 1
            


print("##################")
print("All Done")

## uncomment to save output to file
# f.close()
    
    
