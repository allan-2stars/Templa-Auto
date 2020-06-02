import pyautogui
import pywinauto
import pandas as pd
import time
import csv
from datetime import datetime
from functions.functions_utils import tm_init


## get the appliation handler from the init function


#############################
##
## Site Reassign function
##
#############################
if tm_init() is None:
    print("Can't find Templa on your computer")
else:
    templa = tm_init()[0]
    app = tm_init()[1]
    mainLiteUsertab = templa.child_window(title='LITE Users', control_type='TabItem')
    mainLiteUsertab.click_input()
    mainLiteUserWindow = templa.child_window(title='LITE Users', control_type='Window')

    ########################
    #
    # Setup Excel Sheet
    #
    ########################
    site_reallocate_sheet = 'Email Changing'
    df = pd.read_excel('test.xlsx', sheet_name=site_reallocate_sheet)
    print("starting...")

    for i in df.index:
        user_name = df['NAME']
        user_email = df['EMAIL']
        csm_code = df['CODE']
        status = df['STATUS']
        email_password = df['PASSWORD']

        if status[i] == "Done":
            print(user_name[i] + " is Done")
            continue
        if status[i] == "Skip":
            print(user_name[i] + " is Skipped")
            continue
        if status[i] == "Stop":
            print("Stop here")
            break

        # click on the Code Edit Box
        mainLiteUserWindow.window(title='Code', control_type='ComboBox').click_input()
        pyautogui.typewrite(csm_code[i])


        csmEmailAlreadyAssgined = mainLiteUserWindow.child_window(title=user_email[i], control_type="DataItem")
        
        
        if csmEmailAlreadyAssgined.exists():  
            print("Email: " + user_email[i])
            print("Already set email to " + site_name[i])
            print("#################################")
            print(" ")
            pyautogui.moveRel(-25, 25) 
            pyautogui.click() # reset the select status

        else:
            print("Email Different, need to change")
            print("New Email: " + user_email[i])
            pyautogui.moveRel(-25, 25) 
            pyautogui.doubleClick() # open the site by double click

            userDetailWindow = app.window(title_re='User Details - *')
            userDetailWindow.wait('exists', timeout=15)

            userDetailWindow.child_window(auto_id="txtEmail", control_type="Edit").click_input()
            pyautogui.dragRel(-200,0)
            pyautogui.typewrite(user_email[i])
            userDetailWindow.child_window(title="Email", auto_id="TabItem Key EMAIL", control_type="TabItem").click_input()
            
            ## type new email address
            userDetailWindow.child_window(auto_id="txtSMTPUser", control_type="Edit").click_input()
            pyautogui.dragRel(-200,0)
            pyautogui.typewrite(user_email[i])

            userDetailWindow.child_window(auto_id="txtSMTPPassword", control_type="Edit").click_input()
            pyautogui.dragRel(-200,0)
            pyautogui.typewrite(email_password[i])

            time.sleep(1.5)
            print(user_email[i] + ": is Done now")
            print("###############################")
            print(" ")












            # ## operate the site details analysis window
            # siteAnalysisWindow = userDetailWindow.child_window(title_re='Site Analysis Detail - *')
            # siteAnalysisWindow.wait('exists', timeout=15)
            # siteAnalysisWindow.window(title='Business analysis', control_type='TabItem').click_input()

            # # change CSM and Tablet Number
            # # Edit 39 = CSM, Edit 43 = Tablet
            # siteAnalysisWindow.Edit39.click_input()
            # # pyautogui.moveRel(60,0)
            # time.sleep(1.5)
            # pyautogui.dragRel(-500,0)
            # pyautogui.typewrite(csm[i])
            # pyautogui.press("tab")
            # time.sleep(1.5)
            # # print("Located now to: " + str(csm[i]))
            # pyautogui.moveRel(500,15)
            # pyautogui.click()
            # # siteAnalysisWindow.Edit43.click_input()
            # pyautogui.dragRel(-500,0)
            # time.sleep(1.5)
            # pyautogui.typewrite(tablet[i])
            # pyautogui.press("tab")
            # time.sleep(1.5)
            # pyautogui.moveRel(300,20)
            # pyautogui.click()
            # pyautogui.press("tab")
            # # press Accept button
            # # Save
            # siteAnalysisWindow.Accept.click_input()
            # time.sleep(1.5)
            # userDetailWindow.Save.click_input()
            # #userDetailWindow.window(title='Save', control_type='Button')[0].click_input()
            # time.sleep(1.5)
            # print(user_email[i] + ": is Done now")
            # print("###############################")
            # print(" ")
    