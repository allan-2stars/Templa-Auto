
import pyautogui
import pandas as pd
import time
from datetime import datetime, timedelta, date
import dateutil.relativedelta as relativedelta
import calendar
import pywinauto
from functions.functions_utils import tm_init
##

## need to active the site and open the qa window
## get the appliation handler from the init function
app = tm_init()[1]
print("Starting...")


next_month_today = date.today() + relativedelta.relativedelta(months=1)
next_month = next_month_today.strftime("%m") # type of String


current_year = datetime.now().year
lastday_next_month = calendar.monthrange(current_year, int(next_month))[1]

dateStartString = '01' + next_month + str(current_year)
dateEndString = str(lastday_next_month) + next_month + str(current_year)
# # open analysis details dialouge window
contractDetailWindow = app.window(title_re='Contract - *')
contractDetailWindow.wait('exists', timeout=15)
## Start a New Version

# contractDetailWindow.window(title='New version').click_input(double=True)

## and confirm you want to start a new version
# pyautogui.PAUSE = 2.5
# pyautogui.typewrite('y') ## equivilent to clicking "yes"
# pyautogui.PAUSE = 3.5

# press the tab of QA
contractDetailWindow.child_window(title='QA', control_type='TabItem').click_input()


########################
#
# Setup Excel Sheet
#
########################
sheetLoader = 'ETH' 
df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)
# print("Reading Excel...")
for i in df.index:
    area = df['AREA']
    title = df['TITLE']
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

    use_month = monthJan[i]   
    if next_month == '02':
        use_month = monthFeb[i]
    if next_month == '03':
        use_month = monthMar[i]
    if next_month == '04':
        use_month = monthApr[i]
    if next_month == '05':
        use_month = monthMay[i]
    if next_month == '06':
        use_month = monthJun[i]
    if next_month == '07':
        use_month = monthJul[i]
    if next_month == '08':
        use_month = monthAug[i]
    if next_month == '09':
        use_month = monthSep[i]
    if next_month == '10':
        use_month = monthOct[i]
    if next_month == '11':
        use_month = monthNov[i]
    if next_month == '12':
        use_month = monthDec[i]
        
    # 'x' marks need to set it up, otherwise no need setup.
    if use_month != "x":
        # print('Skipped: ', completeTitle)
        continue

    if status[i] == "Stop":
        print("Stop here")
        break

    if status[i] == "Done" or status[i] == "Skip":
        print('Done: ', completeTitle)
        continue

    print('now is writing... ')
    print(dateStartString + '-' + dateEndString)
    ## click Add    
    contractDetailWindow.child_window(title="Add", auto_id="btnAddQA", control_type="Button").click_input()

    ## in Contrac QA Window
    contractQAWindow = contractDetailWindow.window(title_re='Contract QA - *')
    contractQAWindow.wait('exists', timeout=15)
    contractQAWindow.child_window(auto_id="datEffectiveFrom", control_type="Edit").click_input()

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
    pyautogui.typewrite(qaTemplate[i])
    pyautogui.press('tab')
    ## the title will auto comes up.
    ## add desired one
    pyautogui.typewrite(completeTitle)
    pyautogui.press('tab')
    pyautogui.typewrite(str(task[i]))
    pyautogui.press('tab')
    contractQAWindow.child_window(title="Any time", control_type="RadioButton").click_input()
    contractQAWindow.child_window(auto_id="datNextQA", control_type="Edit").click_input() # next qa edit box
    pyautogui.typewrite(dateStartString)
    pyautogui.press('tab')

    contractQAWindow.Accept.click_input()
    print("QA done: " + completeTitle)

# contractDetailWindow.window(title='Request approval').click_input()
# pyautogui.PAUSE = 2.5
# pyautogui.typewrite('y') ## equivilent to clicking "yes"

print("All QA completed now")
print("##################")
    
    
