
import calendar
from datetime import datetime, timedelta, date
import dateutil.relativedelta as relativedelta


## this function output a list of start date and end date for filter
#####################################
#
#   parameter: relativeMonthNumber
#   type: number
#   if the number is 1, means the month will be next month from current month
#   if the number is -1, means the month will be previous month from current one.
#
#####################################

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