import os
from pywinauto import Application

import calendar
from datetime import datetime, timedelta, date
import dateutil.relativedelta as relativedelta

def tm_init():
    if (os.path.exists(r"E:\TCMS_LIVE\Client Suite")):
        templa_file = r"E:\TCMS_LIVE\Client Suite\TemplaCMS32.exe"
        app = Application(backend='uia').connect(path=templa_file)
    else:
        print("Can't find Templa on your computer")

    templa = app.window(title='TemplaCMS  -  Contract Management System  --  TJS Services Group Pty Ltd LIVE')
    return [templa, app]




def date_range(relativeMonthNumber):
    month_relative_of_today = date.today() + relativedelta.relativedelta(months=relativeMonthNumber)
    month_number_relative = month_relative_of_today.strftime("%m") # type of String


    current_year = datetime.now().year
    lastday_of_relative_month = calendar.monthrange(current_year, int(month_number_relative))[1]

    date_start_string = '01' + month_number_relative + str(current_year)
    date_end_string = str(lastday_of_relative_month) + month_number_relative + str(current_year)
    print(date_start_string)
    print(date_end_string)

    # return two values in a list
    return [date_start_string, date_end_string]