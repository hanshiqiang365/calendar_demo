#author: hanshiqiang365 (developed by 韩思工作室)

import calendar
from datetime import date

mydate = date.today()
year_calendar_str = calendar.calendar(2020)
print(f"{mydate.year}年的日历图：{year_calendar_str}\n")
