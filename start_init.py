
import pyautogui
from pywinauto.application import Application
import os
import time


        

def start_init():
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').connect(path=templa_file)
    print("Get in Main Window...")
    templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')
    templa.wait("exists", timeout=15)
    ## Click on Favourites Menu
    templa.child_window(title="Favourites", control_type="Group").click_input()

    ### the list of title in 'Favourites' menu
    list_favourites = ['Workflow Manager', 'Device Registration', 'Workflow Paths', \
                      'LITE Users', 'Analysis Codes', 'Sites', 'Contracts', 'Contacts']

    def operate_filter(filter_name):
        filter_window = templa.window(title_re=filter_name)
        # Wait filter comes out
        filter_window.wait('exists', timeout=35)
        filter_window.child_window(title="Default criteria", control_type="Button").click_input()
        print("looking for Save button for, " + filter_name )
        filter_window.child_window(title="Save", \
                        auto_id="[Group : save Tools] Tool : CodedMaintenance_saveandclose - Index : 0 ", \
                        control_type="Button").click_input()
        print(filter_name + " Saved, waiting ...")
        time.sleep(7)

    ## Open Contract
    for list_title in list_favourites:
        time.sleep(1) # wait 1 second for opening the menu
        templa.child_window(title=list_title, control_type="DataItem").click_input()

        ## if the window opened need more filter or details to do, use below conditions path
        if list_title == "Sites":
            operate_filter('Site Filter Detail - *')
        
        if list_title == "Contracts":
            operate_filter('Contract Filter Detail - *')

        if list_title == "Contacts":
            operate_filter('Contact Filter Detail - *')
        