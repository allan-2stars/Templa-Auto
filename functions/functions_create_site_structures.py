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
def Create_Site_Structures():
    if tm_init() is None:
        print("Can't find Templa on your computer")
    else:
        templa = tm_init()[0]
        app = tm_init()[1]
        # start 
        print("Starting...")
        # mainSiteStructuresTab = templa.child_window(title='Site Structures', control_type='TabItem')
        # mainSiteStructuresTab.click_input()
        mainSiteStructuresWindow = templa.child_window(title='Site Structures', control_type='Window')
        ########################
        #
        # Setup Excel Sheet
        #
        ########################
        sheetLoader = 'Create Site Structures' 
        df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
        print("Reading Excel...")
        for i in df.index:
            structureCode = df['STRUCTURE CODE']
            description = df['DESCRIPTION']
            area_items = df['AREAS']
            section_end = df['END SECTION']
            status = df['STATUS']

            if status[i] == "Done":
                print(description[i] + " is Done")
                continue

            if status[i] == "Same Section":
                ## no need print out infomation
                ## undless in debug mode
                # print(description[i] + " is under Same Section")
                continue

            if status[i] == "Skip":
                print(description[i] + " is Skipped")
                continue

            if status[i] == "Stop":
                print("Stop here")
                break

            # click on the Code Edit Box
            #  mainSiteStructuresWindow.window(title='Description', control_type='ComboBox').click_input()
            print("click New button ...")
            new_button = templa.child_window(title="New", auto_id="[Group : row Tools] Tool : list_New - Index : 1 ", control_type="Button")
            new_button.click_input()
            SiteStructureWindow = app.window(title_re='Site Structure*')
            SiteStructureWindow.wait('exists', timeout=25)
            print("Site Structure Window opened...")
            pyautogui.PAUSE = 1.5
            pyautogui.press('tab')
            pyautogui.typewrite(description[i])

            SiteStructureWindow.child_window(title='Service areas', control_type='TabItem').click_input()

            ## loop to input area items
            while not section_end[i]:
                print("more area item exists ...")
                
                print("i inside while loop now is: " + str(i))
                SiteStructureWindow.child_window(title="Add area", control_type="Button").click_input()
                pyautogui.typewrite(str(area_items[i]))
                print("added area: " + str(area_items[i]))
                i = i + 1
            print('no more area item left ... ready to save')
            SiteStructureWindow.Save.click_input()