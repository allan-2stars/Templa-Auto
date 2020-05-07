
import pyautogui
from pywinauto.application import Application
import os
import time
from functions.functions_utils import date_range
from functions.functions_utils import tm_init

def start_init():
    if tm_init() is None:
        print("Can't find Templa on your computer")
    else:
        templa = tm_init()[0]
        templa.wait("exists", timeout=15)
        ## Click on Favourites Menu
        templa.child_window(title="Favourites", control_type="Group").click_input()

        ### the list of title in 'Favourites' menu
        list_favourites = ['Workflow Manager', 'Device Registration', 'Workflow Paths', \
                        'LITE Users', 'Analysis Codes', 'Sites', 'Contracts', 'Contacts', 'QA Forms']


        def operate_filter(filter_name):
            filter_window = templa.window(title_re=filter_name)
            # Wait filter comes out
            filter_window.wait('exists', timeout=35)
            if filter_name == 'QA Filter Detail - *':
                filter_window.child_window(title="QA filtering criteria", control_type='TabItem').click_input()
                filter_window.child_window(auto_id="datScheduledDateFrom", control_type="Edit").click_input()
                # get the list of date, and type them in place
                date_range_list = date_range(-1)
                pyautogui.typewrite(date_range_list[0])
                pyautogui.press("tab")
                pyautogui.typewrite(date_range_list[1])
                pyautogui.press("tab")
            # else:
            #     ## for other filter, click default before save
            #     filter_window.child_window(title="Default criteria", control_type="Button").click_input()
            
            print("looking for Save button for, " + filter_name )
            filter_window.child_window(title="Save", \
                            auto_id="[Group : save Tools] Tool : CodedMaintenance_saveandclose - Index : 0 ", \
                            control_type="Button").click_input()
            print(filter_name + " Saved, waiting ...")

        ## Open Contract
        for list_title in list_favourites:
            time.sleep(7) # wait 1 second for opening the menu
            templa.child_window(title=list_title, control_type="DataItem").click_input()

            ## if the window opened need more filter or details to do, use below conditions path
            if list_title == "Sites":
                operate_filter('Site Filter Detail - *')
            
            if list_title == "Contracts":
                operate_filter('Contract Filter Detail - *')

            if list_title == "Contacts":
                operate_filter('Contact Filter Detail - *')

            if list_title == "QA Forms":
                operate_filter('QA Filter Detail - *')
            