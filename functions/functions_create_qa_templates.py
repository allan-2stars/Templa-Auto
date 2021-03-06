import pyautogui
import pywinauto
import pandas as pd
import time
import csv
from functions.functions_utils import tm_init



#############################
##
## Create QA Templates function
##
#############################
def Create_QA_Templates():

    if tm_init() is None:
        print("Can't find Templa on your computer")
    else:
        templa = tm_init()[0]
        app = tm_init()[1]
        # start 
        print("Starting...")
        mainQATemplateTab = templa.child_window(title='QA Templates', control_type='TabItem')
        mainQATemplateTab.click_input()

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
            sections = df['SECTIONS']
            qa_items = df['QA ITEMS']
            qa_item_groups = df['QA ITEM GROUPS']
            status = df['STATUS']

            if status[i] == "Done":
                print(description[i] + " is Done")
                continue

            if status[i] == "Same Template":
                ## no need print out infomation
                ## unless in debug mode
                # print(description[i] + " is under Same Section")
                continue

            if status[i] == "Skip":
                print(description[i] + " is Skipped")
                continue

            if status[i] == "Stop":
                print("Stop here")
                break


            print("click New button ...")
            new_button = templa.child_window(title="New", auto_id="[Group : row Tools] Tool : list_New - Index : 1 ", control_type="Button")
            new_button.click_input()
            QATemplateWindow = app.window(title_re='QA Template*')
            QATemplateWindow.wait('exists', timeout=25)

            print('New Template Creating ...')
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

            ## Start "QA items" tab section
            ## add Overall criteria
            print('Adding QA Items to every Section ...')
            QATemplateWindow.child_window(title="Overall criteria", auto_id="[Node] 0", control_type="DataItem").click_input()
            QATemplateWindow.child_window(title="Add QA criteria", auto_id="[Group : structure Tools] Tool : Item_AddCriteria - Index : 7 ", control_type="Button").click_input()
            QACriteriaWindow = app.window(title_re='QA Criteria')
            QACriteriaWindow.wait('exists', timeout=25)
            QACriteriaWindow.child_window(title="Description", control_type="ComboBox").click_input()
            pyautogui.typewrite(qa_items[i])
            pyautogui.moveRel(-25, 25) 
            pyautogui.click() # open the site by double click
            QACriteriaWindow.Select.click_input()
            QACriteriaWindow.Close.click_input()
            
            ## add qa items for every section
            add_qa_item_button = QATemplateWindow.child_window(title="Add QA item", auto_id="[Group : structure Tools] Tool : Item_AddItem - Index : 6 ", control_type="Button")
            
            ## Loop to every section here
            while True:
                i = i + 1 # i need +1 because, the qa item from the 2nd line
                section_title = QATemplateWindow.child_window(title=str(sections[i]), control_type="DataItem")
                section_title.click_input()
                add_qa_item_button.click_input()

                QAItemsWindow = app.window(title_re='QA Items')
                QAItemsWindow.wait('exists', timeout=25)
                QAItemsWindow.child_window(title="Group", control_type="ComboBox").click_input()
                pyautogui.typewrite(str(qa_item_groups[i]))
                
                print('Adding Items under section: ' + str(sections[i]))
                ## Loop to add QA items for each section here
                while True:
                    QAItemsWindow.child_window(title="Description", control_type="ComboBox").click_input()
                    pyautogui.typewrite(str(qa_items[i]))
                    pyautogui.moveRel(-25, 25) 
                    pyautogui.click() # open the site by double click
                    QAItemsWindow.Select.click_input()
                    print('Added item: ' + qa_items[i])
                    ## check if next items are in the same seciont,
                    ## if not, jump out of the loop
                    try: ## check if next line exists in spreadsheet
                        next_section = sections[i+1]
                    except: ## not exist, then must be the last item already
                        print('last line, no more section !')
                        print('---------------------------')
                        QAItemsWindow.Close.click_input()
                        break
                    else:
                        ## if next line is not the same section, 
                        ## jump out of the while loop, need to add to a new section
                        if sections[i] != sections[i+1]:
                            QAItemsWindow.Close.click_input()
                            break  # note, once break i will not plus one
                    ## next line still int the same section, Go to next line
                    i = i + 1

                QATemplateWindow.wait('exists', timeout=10)
                section_title.click_input(button='left', double=True) ## to collapse items save screen space

                try: ## check if next line exists in spreadsheet
                    next_template = description[i+1]
                except: ## not exist, then must be the last item already
                    print('no more template need to be created, exit !')
                    QATemplateWindow.Save.click_input()
                    break
                else:
                    if description[i] != description[i+1]:
                        QATemplateWindow.Save.click_input()
                        break
                
