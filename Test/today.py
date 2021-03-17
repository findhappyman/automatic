import subprocess
import datetime
import time
#import chinese_calendar
#from chinese_calendar import is_workday
#import pySAP

One_day = 60*60*24*1

print(time.strftime('%d.%m.%Y',time.localtime(time.time()- One_day)))


