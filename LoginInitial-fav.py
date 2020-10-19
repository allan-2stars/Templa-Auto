import pyautogui
from pywinauto.application import Application
import os
import time
from functions.functions_init import start_init 

if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
    templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
    app = Application(backend='uia').start(templa_file)
    #app = Application(backend='uia').connect(path=templa_file)
else:
    print("Can't find Templa on your computer")


loginPage = app['TemplaCMS  -  Login']

loginPage['Edit'].click_input()
pyautogui.typewrite('awa')
loginPage['PasswordEdit'].click_input()
pyautogui.typewrite('wlnce')
loginPage['LoginButton'].click_input()


# Error Active User Exist
errorWindow = loginPage.window(title_re="Existing*")
time.sleep(2)

print("Starting...")
if errorWindow.exists():
    print("Active user exists...")
    # errorWindow.print_control_identifiers()
    redCross_DataItem_1 = errorWindow.child_window(auto_id="0", control_type="DataItem")
    redCross_DataItem_2 = errorWindow.child_window(auto_id="1", control_type="DataItem")
    redCross_DataItem_3 = errorWindow.child_window(auto_id="2", control_type="DataItem")
    # if updated again, increase the auto_id by 1, last time was 51
    redCross_1 = redCross_DataItem_1.child_window(auto_id="52", control_type="Edit")
    redCross_2 = redCross_DataItem_2.child_window(auto_id="52", control_type="Edit")
    redCross_3 = redCross_DataItem_3.child_window(auto_id="52", control_type="Edit")

    print ("red cross 1 existing ?", redCross_1.exists())
    print ("red cross 2 existing ?", redCross_2.exists())
    print ("red cross 3 existing ?", redCross_3.exists())

    if redCross_DataItem_1.exists():
        redCross_1.click_input()
        pyautogui.press('y')
    if redCross_DataItem_2.exists():
        redCross_2.click_input()
        pyautogui.press('y')
    if redCross_DataItem_3.exists():
        redCross_3.click_input()
        pyautogui.press('y')

    # if not (redCross_DataItem_1.exists() or redCross_DataItem_2.exists() or redCross_DataItem_3.exists()):
    #     print("Cannot close the previouse session, red cross button code changed")
    #     print("please use below method to find the correct code")
    #     print("uncommon the function '# errorWindow.print_control_identifiers()'")
    #     print("press any key to exit!")
    #     input()
    
    errorWindow.Continue.click_input()
    start_init()
else:
    start_init()    



