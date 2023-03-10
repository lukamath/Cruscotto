# Importing pandas
import os
import sys
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
import matplotlib as mpl
import matplotlib.pyplot as plt

#scrittura corretta della fascia
now = datetime.now()	
today=now.strftime("%d/%m/%Y")
later=now+timedelta(hours=1)
current_time = now.strftime("%H:%M:%S")
later_time=later.strftime("%H:%M:%S")
timeframe=now.strftime("%H")+ ":00-" + str(later.strftime("%H")) +":00"
hour=now.strftime("%H")
inthour=int(hour)
#print("date and time:",today)

forecast = pd.read_excel('data_source/forecast.xlsx', sheet_name="x")
dayforecast = forecast[['fcst_name', 'Gestione CLT','timeframe',today]].copy()
dayforecast["hour"]=dayforecast["timeframe"].str[:2].astype(int)
dayforecast.to_excel("service/dayforecast.xlsx")
# currentforecast e' il mio forecast fino all'ora corrente
currentforecast=dayforecast.loc[dayforecast["hour"]<=inthour]	
#currentforecast.to_excel("service/currentforecast.xlsx")

listacode = pd.read_excel('data_source/forecast.xlsx', sheet_name="Legenda Inbound")
themap=pd.read_excel('data_source/map.xlsx', sheet_name="Map")
minimap=themap.groupby(['fcst_name', 'Report Activity'], as_index=False).sum()
#print(minimap)
fcstmap=minimap[['fcst_name', 'Report Activity']].copy()

currentforecast=currentforecast.merge(fcstmap, on='fcst_name')
currentforecast.to_excel("service/currentforecast.xlsx")

url = "data_source/code.html"
table = pd.read_html(url)[0]
table.to_excel("data_source/code.xlsx")	
code = pd.read_excel('data_source/code.xlsx',skiprows=2)
code.rename(columns={"&nbsp": "vag"}, inplace=True)
# print('*** Code ***')
# print(code)

#match=listacode.merge(code, left_on='VAG Instradamento',right_on='vag')
match=themap.merge(code, left_on='Coda',right_on='vag')
#match.to_excel("service/match.xlsx")	

match.rename(columns={"Riclassifica": "fcst_name"}, inplace=True)
match.to_excel("service/match&timeframe.xlsx")

currentforecastonname=currentforecast.groupby('fcst_name').sum()
currentforecastonactivity=currentforecast.groupby('Report Activity').sum()
currentforecastonname.to_excel("service/currentforecastonname.xlsx")
currentforecastonactivity.to_excel("service/currentforecastonactivity.xlsx")

matchonfcst=match.groupby('fcst_name').sum()
matchonactivity=match.groupby('Report Activity').sum()
matchonfcst.to_excel("service/matchonfcst.xlsx")
matchonactivity.to_excel("service/matchonactivity.xlsx")

reporthouronfcst=matchonfcst.merge(currentforecastonname,left_on='fcst_name',right_on='fcst_name')
reporthouronfcst['Delta_Offerto']=(reporthouronfcst['Offerte']/reporthouronfcst[today] - 1)*100
reporthouronfcst=reporthouronfcst.drop([1, 'T.A.','Livello di Servizio %','% Cleared','hour'], axis=1)
reporthouronfcst.to_excel('output/reporthouronfcst.xlsx')

print(reporthouronfcst)

reporthouronfcst = reporthouronfcst.cumsum()

#reporthouronfcst.plot();
reporthouronfcst["Offerte"].plot(kind = 'hist')

plt.show()