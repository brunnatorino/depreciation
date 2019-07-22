import pandas as pd
import numpy as np
import matplotlib as mlt
import math
from datetime import datetime, date
import datetime as dt
from datetime import datetime, timedelta
import time
from datetime import date

## this function imports the excel files, change names within read_excel("") to import different files
countries = pd.read_excel('countries1.xlsx')
assets = pd.read_excel('export-2.xlsx', names = ['Acquisition Date','Asset Description','Asset Class','Value','Retirement','Beginning','Closing'])
## prints the data set

assets = assets.query('Value > 0')

countries_map = countries[['Asset Class','Asset']]
countries_map = countries_map.set_index('Asset Class').to_dict()['Asset']
assets['Asset Class'] = assets['Asset Class'].map(countries_map)

## creates three new dataframes to separate countries
dfA = assets.copy()
dfB = assets.copy()
dfC = assets.copy()
## fills last column country with A,B or C --> replace with country abbreviations
dfA['Country'] = 'Germany'
dfB['Country'] = 'B'
dfC['Country'] = 'C'
## setting up for threshold drop
dfA['Threshold'] = dfA['Country']
dfB['Threshold'] = dfB['Country']
dfC['Threshold'] = dfC['Country']
## creating dictionary for mapping
threshold = countries[['Country','Threshold']]
threshold = threshold.set_index('Country').to_dict()['Threshold']
## mapping threshold according to country
dfA['Threshold'] = dfA['Threshold'].map(threshold)
dfB['Threshold'] = dfB['Threshold'].map(threshold)
dfC['Threshold'] = dfC['Threshold'].map(threshold)
## drop values smaller than threshold 
dfA = dfA.query('Value >= Threshold')
dfB = dfB.query('Value >= Threshold')
dfC = dfC.query('Value >= Threshold')
## creating separate dataframes according to country for mapping of useful life
df1 = countries[(countries['Country'] == 'Germany')]
df2 = countries[(countries['Country'] == 'B')]
df3 = countries[(countries['Country'] == 'C')]
## setting up for dictionary
useful_life1 = df1[['Asset','Useful Life']]
useful_life2 = df2[['Asset','Useful Life']]
useful_life3 = df3[['Asset','Useful Life']]
## dictionary for mapping of useful life
useful_life1 = useful_life1.set_index('Asset').to_dict()['Useful Life']
useful_life2 = useful_life2.set_index('Asset').to_dict()['Useful Life']
useful_life3 = useful_life3.set_index('Asset').to_dict()['Useful Life']
## life = type column for mapping
dfA['Life'] = dfA['Asset Class']
dfB['Life'] = dfB['Asset Class']
dfC['Life'] = dfC['Asset Class']
## mapping for useful life
dfA['Life'] = dfA['Life'].map(useful_life1)
dfB['Life'] = dfB['Life'].map(useful_life2)
dfC['Life'] = dfC['Life'].map(useful_life3)


print(dfA.head())

##rounding values with truncate funcion

today = datetime.today()
month = today.month
year = today.year
dfA['Life_In_Months'] = dfA['Life'].mul(12)

## CALCULATING DEPRECIATION TIME FOR COUNTRY A
dfA['date'] = pd.to_datetime(dfA['Acquisition Date'],format='%d%m%Y')
dfA['year'], dfA['month'], dfA['day'] = dfA['date'].dt.year, dfA['date'].dt.month,dfA['date'].dt.day
dfA['Closing'] = pd.to_datetime(dfA['Closing'] ,format='%d%m%Y')
dfA['Beginning'] = pd.to_datetime(dfA['Beginning'] ,format='%d%m%Y')

## adds a month if the day is bigger than 1, also works for december --> adds a year 
## begins new month in the first day
dfA.loc[dfA["day"] != 1,'First_Date_Month'] = dfA['date'] + pd.offsets.MonthBegin(1)
dfA.loc[dfA["day"] == 1,'First_Date_Month'] = dfA['date']
## calculates how many months it has been since the start of the current financial year
dfA['Months_Past'] = ((dfA['Beginning'] - dfA['First_Date_Month'])/np.timedelta64(1, 'M')).astype(int) + 1

## sets to zero all the assets bought this year// they haven't been depreciated yet 
dfA.loc[dfA["Months_Past"] < 0,'Months_Past'] = 0

## calculates the initial balance of the asset in the start of the year 
dfA['Depreciation_Per_Month'] = dfA['Value'].div(dfA['Life_In_Months'])
dfA['Depreciated_Amount'] = dfA['Depreciation_Per_Month'].mul(dfA['Months_Past'])
dfA['Balance_Start'] = dfA['Value'].sub(dfA['Depreciated_Amount'])
dfA = dfA.query('Balance_Start >= 0')

## depreciation for this year
dfA['Depreciation_This_Year_In_Months'] = dfA['Closing'].dt.month - dfA['First_Date_Month'].dt.month + 1
dfA['Amount_To_Depreciate'] = dfA['Depreciation_This_Year_In_Months'].mul(dfA['Depreciation_Per_Month'])
dfA['End of FY19 - Book Value'] = dfA['Balance_Start'] - dfA['Amount_To_Depreciate']

## calculates months until asset is completely depreciated
dfA['Months_To_Zero'] = dfA['Life_In_Months']- dfA['Months_Past']
## returns correct amount to depreciated if asset's life ends this year
dfA.loc[dfA["Balance_Start"] < dfA['Amount_To_Depreciate'],'Amount_To_Depreciate'] = dfA['Balance_Start']
dfA.loc[dfA["Months_To_Zero"] < 12,'Note'] = 'End of Life in Current fiscal Year'

dfA.loc[dfA["End of FY19 - Book Value"] < 0,'End of FY19 - Book Value'] = 0
print(dfA)


## CALCULATING DEPRECIATION TIME FOR COUNTRY B

dfB['Life_In_Months'] = dfB['Life'].mul(12)

## CALCULATING DEPRECIATION TIME FOR COUNTRY A
dfB['date'] = pd.to_datetime(dfB['Acquisition Date'],format='%d%m%Y')
dfB['year'], dfB['month'], dfB['day'] = dfB['date'].dt.year, dfB['date'].dt.month,dfB['date'].dt.day
dfB['Closing'] = pd.to_datetime(dfB['Closing'] ,format='%d%m%Y')
dfB['Beginning'] = pd.to_datetime(dfB['Beginning'] ,format='%d%m%Y')

## adds a month if the day is bigger than 1, also works for december --> adds a year 
## begins new month in the first day
dfB['First_Date_Month'] = dfB['date'] + pd.offsets.MonthBegin(1)

## calculates how many months it has been since the start of the current financial year
dfB['Months_Past'] = ((dfB['Beginning'] - dfB['First_Date_Month'])/np.timedelta64(1, 'M')).astype(int) + 1

## sets to zero all the assets bought this year// they haven't been depreciated yet 
dfB.loc[dfB["Months_Past"] < 0,'Months_Past'] = 0

## calculates the initial balance of the asset in the start of the year 
dfB['Depreciation_Per_Month'] = dfB['Value'].div(dfB['Life_In_Months'])
dfB['Depreciated_Amount'] = dfB['Depreciation_Per_Month'].mul(dfB['Months_Past'])
dfB['Balance_Start'] = dfB['Value'].sub(dfB['Depreciated_Amount'])
dfB = dfB.query('Balance_Start >= 0')

## depreciation for this year
dfB['Depreciation_This_Year_In_Months'] = dfB['Closing'].dt.month - dfB['First_Date_Month'].dt.month + 1
dfB['Amount_To_Depreciate'] = dfB['Depreciation_This_Year_In_Months'].mul(dfB['Depreciation_Per_Month'])
dfB['End of FY19 - Book Value'] = dfB['Balance_Start'] - dfB['Amount_To_Depreciate']

## calculates months until asset is completely depreciated
dfB['Months_To_Zero'] = dfB['Life_In_Months']- dfB['Months_Past']
## returns correct amount to depreciated if asset's life ends this year
dfB.loc[dfB["Balance_Start"] < dfB['Amount_To_Depreciate'],'Amount_To_Depreciate'] = dfB['Balance_Start']
dfB.loc[dfB["Months_To_Zero"] < 12,'Note'] = 'End of Life in Current fiscal Year'

dfB.loc[dfB["End of FY19 - Book Value"] < 0,'End of FY19 - Book Value'] = 0
print(dfB)

## CALCULATING DEPRECIATION TIME FOR COUNTRY C

dfC['Life_In_Months'] = dfC['Life'].mul(12)


dfC['date'] = pd.to_datetime(dfC['Acquisition Date'],format='%d%m%Y')
dfC['year'], dfC['month'], dfC['day'] = dfC['date'].dt.year, dfC['date'].dt.month,dfC['date'].dt.day
dfC['Closing'] = pd.to_datetime(dfC['Closing'] ,format='%d%m%Y')
dfC['Beginning'] = pd.to_datetime(dfC['Beginning'] ,format='%d%m%Y')

## adds a month if the day is bigger than 1, also works for december --> adds a year 
## begins new month in the first day
dfC['First_Date_Month'] = dfC['date'] + pd.offsets.MonthBegin(1)

## calculates how many months it has been since the start of the current financial year
dfC['Months_Past'] = ((dfC['Beginning'] - dfC['First_Date_Month'])/np.timedelta64(1, 'M')).astype(int) + 1

## sets to zero all the assets bought this year// they haven't been depreciated yet 
dfC.loc[dfC["Months_Past"] < 0,'Months_Past'] = 0

## calculates the initial balance of the asset in the start of the year 
dfC['Depreciation_Per_Month'] = dfC['Value'].div(dfC['Life_In_Months'])
dfC['Depreciated_Amount'] = dfC['Depreciation_Per_Month'].mul(dfC['Months_Past'])
dfC['Balance_Start'] = dfC['Value'].sub(dfC['Depreciated_Amount'])
dfC = dfC.query('Balance_Start >= 0')

## depreciation for this year
dfC['Depreciation_This_Year_In_Months'] = dfC['Closing'].dt.month - dfC['First_Date_Month'].dt.month + 1
dfC['Amount_To_Depreciate'] = dfC['Depreciation_This_Year_In_Months'].mul(dfC['Depreciation_Per_Month'])
dfC['End of FY19 - Book Value'] = dfC['Balance_Start'] - dfC['Amount_To_Depreciate']

## calculates months until asset is completely depreciated
dfC['Months_To_Zero'] = dfC['Life_In_Months']- dfC['Months_Past']
## returns correct amount to depreciated if asset's life ends this year
dfC.loc[dfC["Balance_Start"] < dfC['Amount_To_Depreciate'],'Amount_To_Depreciate'] = dfC['Balance_Start']
dfC.loc[dfC["Months_To_Zero"] < 12,'Note'] = 'End of Life in Current fiscal Year'

dfC.loc[dfC["End of FY19 - Book Value"] < 0,'End of FY19 - Book Value'] = 0
print(dfC)


del dfA['date']
del dfA['day']
del dfA['year']
del dfA['month']
del dfB['date']
del dfB['day']
del dfB['year']
del dfB['month']
del dfC['date']
del dfC['day']
del dfC['month']
del dfC['year']

writer = pd.ExcelWriter("output1-sample.xlsx",
                        engine='xlsxwriter',
                        datetime_format='yyyymmdd',
                        date_format='yyyymmdd')

dfA.to_excel(writer, sheet_name = ('Sheet1'))
dfB.to_excel(writer, sheet_name = ('Sheet2'))
dfC.to_excel(writer, sheet_name = ('Sheet3'))

workbook  = writer.book
worksheet = writer.sheets['Sheet1']
worksheet.set_column('B:C', 20)
writer.save()
