import pyautogui
import pandas as pd
import time
import pywinauto
from datetime import datetime
from functions.functions_utils import tm_init



if tm_init() is None:
    print("Can't find Templa on your computer")
else:
    templa = tm_init()[0]
    app = tm_init()[1]

    ## start 
    templa.child_window(title='Product List', control_type='TabItem').click_input()
    mainProductsWindow = templa.child_window(title='Product List', control_type='Window')

    ########################
    #
    # Setup Excel Sheet
    #
    ########################
    sheetLoader = 'Product Cost'
    df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
    productCode = df['PRODUCT-CODE']
    supplierCode = df['SUPPLIER-CODE']
    supplierCodeRe = df['SUPPLIER-CODE-RE']
    itemName = df['ITEMS']
    cost = df['COST']
    status = df['STATUS']
    prefer = df['PREFERRED']
    for i in df.index:
        
        productCode_string = str(productCode[i])
        supplierCode_string = str(supplierCode[i])
        itemName_string = str(itemName[i])

        if status[i] == "Done":
            print(productCode_string + " is Done")
            continue

        if status[i] == "Skip":
            print(productCode_string + " is Skipped")
            continue

        if status[i] == "Stop":
            print("Stop here")
            break

        # click on the Code/Items name to Edit Box
        if sheetLoader == "Urbanest NT Price":
            mainProductsWindow.window(title='Description', control_type='ComboBox').click_input()
            pyautogui.typewrite(itemName_string)
        else:    
            mainProductsWindow.window(title='Code', control_type='ComboBox').click_input()
            pyautogui.typewrite(productCode_string)

        ## check if product exists before open it
        productItem = mainProductsWindow.child_window(title=productCode_string, control_type="DataItem")
    #   productItem.wait('exists', timeout=5)
        if not productItem.exists():
            pyautogui.press('tab')
            print(productCode_string + ' does NOT Exists,')
            print('Nothing changes, go to Next ...')
            print('--------------------------------')
            print(' ')
            ## go to next directly, no product need to change due to non-exists
            continue

        pyautogui.moveRel(0, 25) 
        pyautogui.doubleClick() # open the site by double click
        print("open products window ...")

        # # open analysis details dialouge window
        # #siteDetailWindow = app.window(title_re='Site Detail - *')
        productDetailWindow = app.window(title_re='Products - *')
        productDetailWindow.wait('exists', timeout=35)

        # Go to supplier, only change the cost price
        productDetailWindow.window(title='Suppliers', control_type='TabItem').click_input()
        supplierEntry = productDetailWindow.child_window(title_re=supplierCodeRe[i])
        if not supplierEntry.exists():
            print('supplier NOT Exists')
            productDetailWindow.Add.click_input()
            # open new supplier detail window
            productSupplierWindow = productDetailWindow.child_window(title_re='Product suppliers - *')
            productSupplierWindow.wait('exists', timeout=35)
            # add supplier name by code
            # the supplier text box is focused by default
            print ("add supplier ...")
            pyautogui.typewrite(supplierCode_string)
            pyautogui.press('tab')
            preferredCheckbox = productSupplierWindow.child_window(auto_id="chkIsPreferredSupplier", control_type="CheckBox")
            isChecked = preferredCheckbox.get_toggle_state()

            # check if match with Excel sheet data
            if str(isChecked) != str(prefer[i]):
                preferredCheckbox.toggle()
            # you can also use tab tab to go down
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(productCode_string)
            pyautogui.press('tab')
            # check prefer checkbox
            # add/change price
            #productSupplierWindow.child_window(auto_id="numUnitCost", control_type="Edit").click_input()
            pyautogui.typewrite(str(cost[i]))
            #pyautogui.press('tab')
        else:  
            print('supplier exists')
            # open specific supplier item
            supplierEntry.click_input(button='left', double=True)
            productSupplierWindow = productDetailWindow.child_window(title_re='Product suppliers - *')
            productSupplierWindow.wait('exists', timeout=35)
            # add/change price
            productSupplierWindow.child_window(auto_id="numUnitCost", control_type="Edit").click_input()
            pyautogui.typewrite(str(cost[i]))
            print('price changed ...')
            pyautogui.press('tab')

        # Save
        productSupplierWindow.Accept.click_input()
        pyautogui.PAUSE = 2.5
        productDetailWindow.Save.click_input()
        time.sleep(5)
        print (productCode_string + ' ' + itemName_string + " is Done now")
        print('#######################')
        print(' ')



    


