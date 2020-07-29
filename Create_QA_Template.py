import pyautogui
import pywinauto
import pandas as pd
import time
import csv
from functions.functions_utils import tm_init


#############################
##
## QA Recipients Reallocation
##
#############################

if tm_init() is None:
    print("Can't find Templa on your computer")
else:
    templa = tm_init()[0]
    app = tm_init()[1]
    # start 
    print("Starting...")
    # mainSiteStructuresTab = templa.child_window(title='Site Structures', control_type='TabItem')
    # mainSiteStructuresTab.click_input()
    mainQATemplateWindow = templa.child_window(title='QA Templates', control_type='Window')

    print("click New button ...")
    new_button = templa.child_window(title="New", auto_id="[Group : row Tools] Tool : list_New - Index : 1 ", control_type="Button")
    new_button.click_input()
    QATemplateWindow = app.window(title_re='QA Template*')
    QATemplateWindow.wait('exists', timeout=25)
    #QATemplateWindow.print_control_identifiers()
########################
    #
    # Setup Excel Sheet
    #
    ########################
    sheetLoader = 'Create QA Templates' 
    df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
    print("Reading Excel...")
    for i in df.index:
        site_structure = df['SITE STRUCTURE']
        description = df['DESCRIPTION']
        override_score_card = df['OVERRIDE SCORE CARD']
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

        description_textbox = QATemplateWindow.child_window(auto_id="txtDescription", control_type="Edit")
        description_textbox.click_input()
        pyautogui.typewrite(description[i])
        time.sleep(1)
        pyautogui.press('tab')
        pyautogui.typewrite("Standard")
        time.sleep(1)
        pyautogui.press('tab')
        pyautogui.typewrite(site_structure[i])
        time.sleep(1)
        pyautogui.press('tab')

        print_score_last_page_checkbox = QATemplateWindow.child_window(auto_id="chkPrintScoreCardOnLastPageOnly", control_type="CheckBox")
        print_score_last_page_status = print_score_last_page_checkbox.get_toggle_state()
        if not print_score_last_page_status:
            print_score_last_page_checkbox.toggle()

        pyautogui.press('tab')
        pyautogui.typewrite(override_score_card[i])

        print_image_checkbox = QATemplateWindow.child_window(auto_id="chkPrintImages", control_type="CheckBox")
        print_image_status = print_image_checkbox.get_toggle_state()
        if not print_image_status:
            print_image_checkbox.toggle()

        auto_email_complete_checkbox = QATemplateWindow.child_window(auto_id="chkAutoEmailOnComplete", control_type="CheckBox")
        auto_email_complete_status = auto_email_complete_checkbox.get_toggle_state()
        if not auto_email_complete_status:
            auto_email_complete_checkbox.toggle()


        QATemplateWindow.child_window(title="QA items", auto_id="TabItem Key items", control_type="TabItem").click_input()