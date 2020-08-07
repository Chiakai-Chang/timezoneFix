# -*- coding: utf-8 -*-
"""
Created on Fri Aug  7 16:43:44 2020

@author: Chiakai
"""

import pandas as pd
import pytz
import datetime

file = 'yahoo_clean_log.xls'

df = pd.read_excel(file)

#時間list跟時區list
timeList = df['Login Time']
timezoneList = df['timezone']

#時區
#pytz.pytz.all_timezones可查時區
TW = pytz.timezone('Asia/Taipei')
GMT = pytz.timezone('GMT')
PST = pytz.timezone('PST8PDT')

#輸出統一格式
timeForm = '%Y-%m-%d %H:%M:%S'

#時區轉換
#先指定時區，如：twdt = TW.localize(now)
#再將時區轉換，如：pstdt = twdt.astimezone(PST)
#再將時區轉換，如：pstdt.strftime('%Y-%m-%d %H:%M:%S')

time_fix = []
for i, d in enumerate(timeList):
    try:
        tempTime = datetime.datetime.strptime(str(d),timeForm)
    except :
        tempTime = datetime.datetime.fromtimestamp(int(str(d)))
    tempTimezone = timezoneList[i] 
    if 'GMT' in tempTimezone:
        gmtdt = GMT.localize(tempTime)
        twdt = gmtdt.astimezone(TW)
        temp_time_fix = twdt.strftime('%Y-%m-%d %H:%M:%S')
        time_fix.append(temp_time_fix)
    elif 'PST' in tempTimezone:
        pstdt = PST.localize(tempTime)
        twdt = pstdt.astimezone(TW)
        temp_time_fix = twdt.strftime('%Y-%m-%d %H:%M:%S')
        time_fix.append(temp_time_fix)
    else:
        time_fix.append(d)
        
#完成，匯入dataframe
df['timeFix'] = time_fix

#輸出xls
df.to_excel('fix.xls')
