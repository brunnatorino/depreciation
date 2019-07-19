import pandas as pd
import numpy as np
import matplotlib as mlt
import math
from datetime import datetime, date

## this function imports the excel files, change names within read_excel("") to import different files
countries = pd.read_excel('example_countries.xlsx')
assets = pd.read_excel('sample_assets.xlsx', names = ['Asset','Value','Acquisition Date'])
## prints the data set
print(assets)
#creates new datasets using only two columns from the original dataframe 
dictassets = assets[['Asset','Value']]
dictassets2 = assets[['Asset','Acquisition Date']]
# creates dictionaries with the same index Asset to map them later with value and acquistion date
dictassets = dictassets.set_index('Asset').to_dict()['Value']
dictassets2 = dictassets2.set_index('Asset').to_dict()['Acquisition Date']
# creates a new column called value in the countries dataset and copies the names of the assets for the mapping 
# same with acquisition date
countries['Value'] = countries['Asset']
countries['Acquisition Date'] = countries['Asset']
# maps the values while dropping those that are not in the dictionary 
countries['Value'] = countries['Value'].map(dictassets)
countries['Acquisition Date'] = countries['Acquisition Date'].map(dictassets2)
# creates three different dataframes for each country, now they ONLY contain one country's information
dfA = countries[(countries['Country'] == 'A')]
dfB = countries[(countries['Country'] == 'B')]
dfC = countries[(countries['Country'] == 'C')]
# drops assets whose value is lower than the threshold (notice that EQUAL TO or LARGER stays)
dfA = dfA.query('Value >= Threshold')
dfB = dfB.query('Value >= Threshold')
dfC = dfC.query('Value >= Threshold')
## print new dataframes for a check
print(dfA)
print(dfB)
print(dfC)

import datetime as dt
from datetime import datetime, timedelta
import time
from datetime import date

today = datetime.today()
month = today.month
year = today.year

dfA['date'] = pd.to_datetime(dfA['Acquisition Date'],format='%d%m%Y')
dfA['year'], dfA['month'], dfA['day'] = dfA['date'].dt.year, dfA['date'].dt.month,dfA['date'].dt.day
dfA['DepreDate'] = dfA['date'].copy()
dfA['Closing'] = pd.to_datetime(dfA['Closing'] ,format='%d%m%Y')

## adds a month if the day is bigger than 1, also works for december --> adds a year 
## begins new month in the first day
dfA['First_Date_Month'] = dfA['date'] + pd.offsets.MonthBegin(1)
dfA['Time_Left_In_Months'] = dfA['Closing'].dt.month - dfA['First_Date_Month'].dt.month

## UP TO 2014

dfA.loc[dfA["First_Date_Month"].dt.year == 2018,'Time_Left_In_Months'] = dfA["Time_Left_In_Months"] + 12
dfA.loc[dfA["First_Date_Month"].dt.year == 2017,'Time_Left_In_Months'] = dfA["Time_Left_In_Months"] + 24
dfA.loc[dfA["First_Date_Month"].dt.year == 2016,'Time_Left_In_Months'] = dfA["Time_Left_In_Months"] + 36
dfA.loc[dfA["First_Date_Month"].dt.year == 2015,'Time_Left_In_Months'] = dfA["Time_Left_In_Months"] + 48
dfA.loc[dfA["First_Date_Month"].dt.year == 2014,'Time_Left_In_Months'] = dfA["Time_Left_In_Months"] + 60

closing = date(today.year, 10, 1)
print('using as closing date:',closing)

print(dfA)

today = datetime.today()
month = today.month
year = today.year

dfB['date'] = pd.to_datetime(dfB['Acquisition Date'],format='%d%m%Y')
dfB['year'], dfB['month'], dfB['day'] = dfB['date'].dt.year, dfB['date'].dt.month,dfB['date'].dt.day
dfB['DepreDate'] = dfB['date'].copy()
dfB['Closing'] = pd.to_datetime(dfB['Closing'] ,format='%d%m%Y')

## adds a month if the day is bigger than 1, also works for december --> adds a year 
dfB['First_Date_Month'] = dfB['date'] + pd.offsets.MonthBegin(1)
dfB['Time_Left_In_Months'] = dfB['Closing'].dt.month - dfB['First_Date_Month'].dt.month

## UP TO 2014

dfB.loc[dfB["First_Date_Month"].dt.year == 2018,'Time_Left_In_Months'] = dfB["Time_Left_In_Months"] + 12
dfB.loc[dfB["First_Date_Month"].dt.year == 2017,'Time_Left_In_Months'] = dfB["Time_Left_In_Months"] + 24
dfB.loc[dfB["First_Date_Month"].dt.year == 2016,'Time_Left_In_Months'] = dfB["Time_Left_In_Months"] + 36
dfB.loc[dfB["First_Date_Month"].dt.year == 2015,'Time_Left_In_Months'] = dfB["Time_Left_In_Months"] + 48
dfB.loc[dfB["First_Date_Month"].dt.year == 2014,'Time_Left_In_Months'] = dfB["Time_Left_In_Months"] + 60

closing = date(today.year, 10, 1)
print('using as closing date:',closing)

print(dfB)

today = datetime.today()
month = today.month
year = today.year

dfC['date'] = pd.to_datetime(dfC['Acquisition Date'],format='%d%m%Y')
dfC['year'], dfC['month'], dfC['day'] = dfC['date'].dt.year, dfC['date'].dt.month,dfC['date'].dt.day
dfC['DepreDate'] = dfC['date'].copy()
dfC['Closing'] = pd.to_datetime(dfC['Closing'] ,format='%d%m%Y')

## adds a month if the day is bigger than 1, also works for december --> adds a year 
dfC['First_Date_Month'] = dfC['date'] + pd.offsets.MonthBegin(1)
dfC['Time_Left_In_Months'] = dfC['Closing'].dt.month - dfC['First_Date_Month'].dt.month

## UP TO 2014

dfC.loc[dfC["First_Date_Month"].dt.year == 2018,'Time_Left_In_Months'] = dfC["Time_Left_In_Months"] + 12
dfC.loc[dfC["First_Date_Month"].dt.year == 2017,'Time_Left_In_Months'] = dfC["Time_Left_In_Months"] + 24
dfC.loc[dfC["First_Date_Month"].dt.year == 2016,'Time_Left_In_Months'] = dfC["Time_Left_In_Months"] + 36
dfC.loc[dfC["First_Date_Month"].dt.year == 2015,'Time_Left_In_Months'] = dfC["Time_Left_In_Months"] + 48
dfC.loc[dfC["First_Date_Month"].dt.year == 2014,'Time_Left_In_Months'] = dfC["Time_Left_In_Months"] + 60

closing = date(today.year, 10, 1)
print('using as closing date:',closing)

print(dfC)
