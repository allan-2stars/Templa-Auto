import pyautogui
from pywinauto.application import Application
import os
import time
from start_init import start_init 

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
# errorWindow.print_control_identifiers()
time.sleep(5)
#errorWindow.wait("exists",timeout=15)
print("Starting...")
if errorWindow.exists():
    print("Active user exists...")
    redCross = errorWindow.child_window(auto_id="50", control_type="Edit")
    print ("existing ?", redCross.exists())
    if redCross.exists():
        redCross.click_input()
        pyautogui.press('y')
        errorWindow.Continue.click_input()
        start_init()
    else:
        print("Cannot close the previouse session, red cross button code changed")
        print("please use below method to find the correct code")
        print("uncommon the function '# errorWindow.print_control_identifiers()'")
        print("exiting ... bye!")
else:
    start_init()    



