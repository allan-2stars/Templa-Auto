import pyautogui
import pandas as pd
import time
import pywinauto
from datetime import datetime
from functions.functions_utils import tm_init

## get the appliation handler from the init function
templa = tm_init()[0]
app = tm_init()[1]

print("Starting...")
print("Site Activated")



########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'TEH' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
print("Reading Excel...")
for i in df.index:
    area = df['AREA']
    title = df['TITLE']
    dateStart = df['DATE START']
    dateEND = df['DATE END']
    qaTemplate = df['QA TEMPLATE']
    task = df['TASK']
    status = df['STATUS']
    monthJan = df['Jan']
    monthFeb = df['Feb']
    monthMar = df['Mar']
    monthApr = df['Apr']
    monthMay = df['May']
    monthJun = df['Jun']
    monthJul = df['Jul']
    monthAug = df['Aug']
    monthSep = df['Sep']
    monthOct = df['Oct']
    monthNov = df['Nov']
    monthDec = df['Dec']
    completeTitle = area[i] + ' - ' + title[i]

    print(completeTitle)
    # 'x' marks need to set it up, otherwise no need setup.
    ## cleprint(completeTitle + ' - ' + monthFeb)
    if monthAug[i] != "x":
       # print("No Need Setup ...")
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    if status[i] == "Done":
        print(completeTitle + " is Done")
       
        continue




    ## click Add    
    # contractDetailWindow.child_window(title="Add", auto_id="btnAddQA", control_type="Button").click_input()


    ## in Contrac QA Window
    # contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
    # contractQAWindow.wait('exists', timeout=15)
    # contractQAWindow.child_window(auto_id="datEffectiveFrom", control_type="Edit").click_input()
    pyautogui.press('tab')
    dateStartString = "01082019"
    dateEndString = "31082019"
    pyautogui.typewrite(dateStartString)
    pyautogui.press('tab')
    pyautogui.typewrite(dateEndString)
    #nextDateString = str(nextQaDate[i])

    # get the date character one by one and type in
    # for letter in nextDateString:
    #     pyautogui.typewrite(letter)
    pyautogui.press('tab')

    ## contractQAWindow.child_window(auto_id="datEffectiveTo", control_type="Edit")
    pyautogui.press('tab')
    
    time.sleep(1)
    pyautogui.typewrite(qaTemplate[i])
    pyautogui.press('tab')
    ## the title will auto comes up.
    ## add desired one
    time.sleep(1)
    pyautogui.typewrite(completeTitle)
    pyautogui.press('tab')
    pyautogui.typewrite(str(task[i]))
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(1)

    # contractQAWindow.child_window(title="Any time", control_type="RadioButton").click_input()
    # contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box
    pyautogui.typewrite(dateStartString)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')


    # contractQAWindow.Accept.click_input()
    print("QA done: " + completeTitle)

# contractDetailWindow.window(title='Request approval').click_input()
# pyautogui.PAUSE = 2.5
# pyautogui.typewrite('y') ## equivilent to clicking "yes"

print("All QA completed now")
print("##################")
    
    
